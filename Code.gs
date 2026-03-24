function doGet(e) {
  const teacherParam = (e.parameter.teacher || '').trim();
  const classParam = (e.parameter.class || '').trim();

  const template = HtmlService.createTemplateFromFile('Index');
  template.teacher = teacherParam;
  template.className = classParam;

  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Class Quiz Completion Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTeacherSheetNames() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const ignored = ['Roster', 'Lookups', 'emailOverrides', 'Y10BTEC', 'ClassroomIDs'];
  return sheets
    .map(sheet => sheet.getName())
    .filter(name => !ignored.includes(name));
}

function getClassesForTeacher(teacherInitials) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(teacherInitials);
  if (!sheet) throw new Error(`Sheet "${teacherInitials}" not found`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues(); // A, B, C
  const classes = data
    .filter(row => row[0] && row[1] && row[2])
    .map(row => ({
      shortName: String(row[0]).trim(),
      fullName: String(row[1]).trim(),
      classroomId: String(row[2]).trim()
    }));

  const cleaned = classes.filter(row => !isInvalidClassroomId_(row.classroomId));

  Logger.log(
    '[getClassesForTeacher] teacher=%s rows=%s returned=%s removedInvalid=%s sample=%s',
    teacherInitials,
    data.length,
    cleaned.length,
    classes.length - cleaned.length,
    JSON.stringify(cleaned.slice(0, 10))
  );

  return cleaned;
}


function logSelectedClassRow(teacherInitials, selected) {
  Logger.log(
    '[classSelection] teacher=%s short=%s full=%s classroomId=%s',
    String(teacherInitials || ''),
    String((selected && selected.shortName) || ''),
    String((selected && selected.fullName) || ''),
    String((selected && selected.classroomId) || '')
  );
  return true;
}


function getLatestFormResponders(classId, rosterEmails) {
  try {
    const course = Classroom.Courses.get(classId);
    const folderId = course.teacherFolder?.id;

    if (!folderId) return { error: "No teacher folder linked to this course." };

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_FORMS);

    let latestForm = null;
    while (files.hasNext()) {
      const file = files.next();
      if (!latestForm || file.getDateCreated() > latestForm.getDateCreated()) {
        latestForm = file;
      }
    }

    if (!latestForm) return { error: "No forms found in teacher folder." };

    const form = FormApp.openById(latestForm.getId());
    const responses = form.getResponses();

    const latestByEmail = {};
    for (let i = responses.length - 1; i >= 0; i--) {
      const email = responses[i].getRespondentEmail()?.toLowerCase().trim();
      if (!email || latestByEmail[email]) continue;
      latestByEmail[email] = true;
    }

    const responderData = rosterEmails.map(student => {
      const email = student.email.toLowerCase().trim();
      const fullName = student.fullName;
      const completed = !!latestByEmail[email];

      return { email, name: fullName, completed };
    });

    return {
      responderData,
      formId: latestForm.getId(),
      formName: latestForm.getName()
    };

  } catch (err) {
    return { error: "Failed to get form responders: " + err.message };
  }
}


function getClassRoster(fullClassName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
  if (!sheet) throw new Error('Roster sheet not found');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Read C–F so we can filter by class (C) and load D, E, F as requested
  const data = sheet.getRange(2, 3, lastRow - 1, 4).getValues(); // C–F

  const CLASS_COL_IDX   = 0; // C
  const FIRST_COL_IDX   = 1; // D
  const SURNAME_COL_IDX = 2; // E
  const EMAIL_COL_IDX   = 3; // F

  const matchingRows = data.filter(row => String(row[CLASS_COL_IDX]).trim() === fullClassName);

  const nameGroups = new Map();

  for (const row of matchingRows) {
    const firstName = String(row[FIRST_COL_IDX]).trim();
    const surname   = String(row[SURNAME_COL_IDX]).trim();
    const email     = String(row[EMAIL_COL_IDX]).trim();
    if (!firstName) continue;

    const key = firstName.toLowerCase();
    if (!nameGroups.has(key)) nameGroups.set(key, []);
    nameGroups.get(key).push({ email, firstName, surname });
  }

  const disambiguated = [];

  for (const [, group] of nameGroups.entries()) {
    if (group.length === 1) {
      const entry = group[0];
      disambiguated.push({ fullName: entry.firstName, email: entry.email });
      continue;
    }

    let suffixLength = 0;
    let finalNames = [];

    while (true) {
      const uniqueNames = new Set();
      finalNames = [];
      let allUnique = true;

      for (const entry of group) {
        const suffix = entry.surname.substring(0, suffixLength + 1);
        const name = suffix ? `${entry.firstName} ${suffix}` : entry.firstName;

        if (uniqueNames.has(name)) {
          allUnique = false;
          break;
        }

        uniqueNames.add(name);
        finalNames.push({ fullName: name, email: entry.email });
      }

      if (allUnique) break;
      suffixLength++;

      // If we've exhausted surnames, stop and keep just the first
      if (suffixLength >= Math.max(...group.map(e => e.surname.length))) {
        finalNames = [finalNames[0]];
        break;
      }
    }

    disambiguated.push(...finalNames);
  }

  return disambiguated;
}



function _getClassRoster(fullClassName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
  if (!sheet) throw new Error(`Roster sheet not found`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 2, lastRow - 1, 4).getValues(); // B–E
  const matchingRows = data.filter(row => String(row[1]).trim() === fullClassName);

  const NAME_COL_IDX = 2; // D
  const EMAIL_COL_IDX = 0; // B
  const SURNAME_COL_IDX = 3; // E

  const nameGroups = new Map();

  matchingRows.forEach(row => {
    const email = String(row[EMAIL_COL_IDX]).trim();
    const firstName = String(row[NAME_COL_IDX]).trim();
    const surname = String(row[SURNAME_COL_IDX]).trim();

    const key = firstName.toLowerCase();
    if (!nameGroups.has(key)) nameGroups.set(key, []);

    nameGroups.get(key).push({ email, firstName, surname });
  });

  const disambiguated = [];

for (const [firstNameKey, group] of nameGroups.entries()) {
  if (group.length === 1) {
    const entry = group[0];
    disambiguated.push({ fullName: entry.firstName, email: entry.email });
    continue;
  }

  let suffixLength = 0;
  let uniqueNames = new Set();
  let finalNames = [];

  while (true) {
    uniqueNames.clear();
    finalNames = [];
    let allUnique = true;

    for (const entry of group) {
      const suffix = entry.surname.substring(0, suffixLength + 1);
      const name = entry.firstName + (suffix ? ' ' + suffix : '');

      if (uniqueNames.has(name)) {
        allUnique = false;
        break;
      }

      uniqueNames.add(name);
      finalNames.push({ fullName: name, email: entry.email });
    }

    if (allUnique) break;
    suffixLength++;
    
    // If we've exhausted surnames, stop and keep just the first
    if (suffixLength >= Math.max(...group.map(e => e.surname.length))) {
      finalNames = [finalNames[0]]; // Only keep one
      break;
    }
  }

  disambiguated.push(...finalNames);
}


  return disambiguated;
}

function getRespondersByFormId(formId, roster) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `matched_${formId}`;
    const matchedSet = new Set(JSON.parse(cache.get(cacheKey) || '[]'));

    const form = FormApp.openById(formId);
    const latest = {};
    form.getResponses().reverse().forEach(r => {
      const e = r.getRespondentEmail();
      if (e) latest[e.toLowerCase().trim()] = true;
    });
    const formEmails = Object.keys(latest);
    const rosterEmails = roster.map(s => s.email.toLowerCase().trim());
    const responderData = [];
    const newlyMatched = [];

    roster.forEach(stu => {
      const email = stu.email.toLowerCase().trim();
      const matched = !!latest[email];
      responderData.push({
        email,
        name: stu.fullName,
        completed: matched
      });
      if (matched && !matchedSet.has(email)) {
        newlyMatched.push(email);
      }
    });

    if (newlyMatched.length) {
      newlyMatched.forEach(e => matchedSet.add(e));
      cache.put(cacheKey, JSON.stringify(Array.from(matchedSet)), 21600); // 6 h TTL
    }

    return responderData;

  } catch (err) {
    return { error: 'Failed to get responders: ' + err.message };
  }
}

function diagnoseMismatch(a, b) {
  if (a.split('@')[0] === b.split('@')[0]) return 'domain differs';
  if (a.replace(/[^a-z0-9]/g, '') === b.replace(/[^a-z0-9]/g, ''))
    return 'punctuation / case only';
  return '';
}

function getCourseAssignments(classId) {
  try {
    const resolvedClassId = resolveCourseId_(classId);
    const response = Classroom.Courses.CourseWork.list(resolvedClassId, {
      orderBy: 'updateTime desc',
      pageSize: 200
    });

    const assignments = (response.courseWork || [])
      .filter(work => work.workType === 'ASSIGNMENT')
      .map(work => ({
        id: String(work.id),
        title: work.title || 'Untitled assignment'
      }));

    return assignments;
  } catch (err) {
    throw new Error('Failed to load assignments: ' + err.message);
  }
}

function getStudentStatusesForAssignments(classId, roster, spellingAssignmentId, classroomAssignmentId) {
  try {
    const resolvedClassId = resolveCourseId_(classId);
    const spellingStatuses = getStatusesByUserId_(resolvedClassId, spellingAssignmentId);
    const classroomStatuses = getStatusesByUserId_(resolvedClassId, classroomAssignmentId);

    const students = Classroom.Courses.Students.list(resolvedClassId).students || [];
    const userIdByEmail = new Map(
      students
        .filter(s => s.profile && s.profile.emailAddress)
        .map(s => [s.profile.emailAddress.toLowerCase().trim(), String(s.userId)])
    );

    return roster.map(student => {
      const email = String(student.email || '').toLowerCase().trim();
      const userId = userIdByEmail.get(email);

      return {
        email,
        name: student.fullName,
        spellingCompleted: userId ? !!spellingStatuses.get(userId) : false,
        classroomCompleted: userId ? !!classroomStatuses.get(userId) : false
      };
    });
  } catch (err) {
    throw new Error('Failed to load student statuses: ' + err.message);
  }
}


function resolveCourseId_(rawClassId) {
  const classId = String(rawClassId || '').trim();
  if (!classId) throw new Error('Classroom ID is empty');

  if (isInvalidClassroomId_(classId)) {
    throw new Error(`Classroom ID is invalid placeholder: ${classId}`);
  }

  const candidates = [classId];

  const numericFromUrl = classId.match(/courses\/(\d+)/i);
  if (numericFromUrl) candidates.push(numericFromUrl[1]);

  const tokenFromUrl = classId.match(/\/c\/([A-Za-z0-9_-]+)/i);
  if (tokenFromUrl) candidates.push(tokenFromUrl[1]);

  for (const candidate of Array.from(new Set(candidates))) {
    try {
      const course = Classroom.Courses.get(candidate);
      if (course && course.id) return String(course.id);
    } catch (err) {
      // continue to fallback attempts
    }
  }

  const matchedByScan = findCourseIdByScan_(candidates);
  if (matchedByScan) return matchedByScan;

  Logger.log('[resolveCourseId_] failed raw=%s candidates=%s', classId, JSON.stringify(candidates));
  throw new Error(`Classroom course not found for ID/value: ${classId}`);
}


function isInvalidClassroomId_(value) {
  const normalized = String(value || '').trim().toUpperCase();
  return normalized === '#N/A'
    || normalized === '#REF!'
    || normalized === '#VALUE!'
    || normalized === 'N/A'
    || normalized === 'NA';
}


function findCourseIdByScan_(candidates) {
  const values = new Set(candidates.map(v => String(v || '').trim()).filter(Boolean));
  let pageToken;

  do {
    const response = Classroom.Courses.list({
      teacherId: 'me',
      pageSize: 100,
      pageToken: pageToken
    });

    const courses = response.courses || [];
    for (const course of courses) {
      const id = String(course.id || '');
      const alternateLink = String(course.alternateLink || '');
      const enrollmentCode = String(course.enrollmentCode || '');

      if (values.has(id)) return id;
      if (values.has(enrollmentCode)) return id;

      for (const value of values) {
        if (value && alternateLink && alternateLink.indexOf(value) !== -1) {
          return id;
        }
      }
    }

    pageToken = response.nextPageToken;
  } while (pageToken);

  return null;
}

function getStatusesByUserId_(classId, courseWorkId) {
  const statuses = new Map();
  if (!courseWorkId) return statuses;

  let pageToken;
  do {
    const response = Classroom.Courses.CourseWork.StudentSubmissions.list(classId, courseWorkId, {
      pageSize: 200,
      pageToken: pageToken
    });

    const submissions = response.studentSubmissions || [];
    submissions.forEach(submission => {
      const userId = String(submission.userId);
      statuses.set(userId, isSubmissionCompleted_(submission));
    });

    pageToken = response.nextPageToken;
  } while (pageToken);

  return statuses;
}

function isSubmissionCompleted_(submission) {
  const hasAssignedGrade = submission.assignedGrade !== undefined && submission.assignedGrade !== null;
  const hasDraftGrade = submission.draftGrade !== undefined && submission.draftGrade !== null;
  return hasAssignedGrade || hasDraftGrade;
}
