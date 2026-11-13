/************************************************************
 * SMART EXAM
 * Sistem Ujian Online & Manajemen Evaluasi Sekolah
 * Backend Google Apps Script + Google Sheets
 *
 * Roles:
 * - Super Admin
 * - Guru
 * - Siswa
 *
 * API:
 * api(data)
 ************************************************************/

const APP = {
  NAME: 'SMART EXAM',
  VERSION: '1.0.0',
  TIMEZONE: 'Asia/Jakarta',
  SESSION_HOURS: 12,
  SHEETS: {
    USERS: 'Users',
    EXAMS: 'Exams',
    QUESTIONS: 'Questions',
    RESULTS: 'Results',
    ANSWERS: 'Answers',
    SESSIONS: 'Sessions',
    LOGS: 'ActivityLogs',
    SETTINGS: 'SchoolSettings',
    VIOLATIONS: 'Violations',
    ATTEMPTS: 'ExamAttempts'
  }
};

const HEADERS = {
  Users: [
    'id',
    'username',
    'passwordHash',
    'salt',
    'name',
    'email',
    'role',
    'className',
    'active',
    'createdAt'
  ],

  Exams: [
    'id',
    'title',
    'subject',
    'className',
    'description',
    'duration',
    'token',
    'startAt',
    'endAt',
    'shuffleQuestions',
    'shuffleAnswers',
    'status',
    'teacherId',
    'createdAt'
  ],

  Questions: [
    'id',
    'examId',
    'question',
    'type',
    'options',
    'answerKey',
    'weight',
    'createdAt'
  ],

  Results: [
    'id',
    'examId',
    'studentId',
    'studentName',
    'className',
    'score',
    'correct',
    'wrong',
    'total',
    'submittedAt'
  ],

  Answers: [
    'id',
    'resultId',
    'examId',
    'studentId',
    'questionId',
    'answer',
    'isCorrect',
    'score',
    'createdAt'
  ],

  Sessions: [
    'token',
    'userId',
    'role',
    'expiresAt',
    'createdAt'
  ],

  ActivityLogs: [
    'id',
    'userId',
    'username',
    'activity',
    'details',
    'createdAt'
  ],

  SchoolSettings: [
    'key',
    'value'
  ],

  Violations: [
    'id',
    'examId',
    'studentId',
    'type',
    'details',
    'createdAt'
  ],

  ExamAttempts: [
    'id',
    'examId',
    'studentId',
    'startedAt',
    'expiresAt',
    'status',
    'submittedAt'
  ]
};


/************************************************************
 * WEB APP
 ************************************************************/

function doGet() {
  setupDatabase();

  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


/************************************************************
 * DATABASE
 ************************************************************/

function setupDatabase() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDatabase_();

    Object.keys(HEADERS).forEach(function(name) {
      let sheet = ss.getSheetByName(name);

      if (!sheet) {
        sheet = ss.insertSheet(name);
      }

      const headers = HEADERS[name];

      if (sheet.getLastRow() === 0) {
        sheet.getRange(
          1,
          1,
          1,
          headers.length
        ).setValues([headers]);

        sheet.setFrozenRows(1);
        sheet.getRange(
          1,
          1,
          1,
          headers.length
        ).setFontWeight('bold');
      }
    });

    seedDefaultData_(ss);

    return {
      status: 'success',
      message: 'Database SMART EXAM berhasil disiapkan.'
    };

  } finally {
    lock.releaseLock();
  }
}


function getDatabase_() {
  const properties =
    PropertiesService.getScriptProperties();

  const databaseId =
    properties.getProperty('SMART_EXAM_DATABASE_ID');

  if (databaseId) {
    return SpreadsheetApp.openById(databaseId);
  }

  const active =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!active) {
    throw new Error(
      'Spreadsheet belum terhubung. Jalankan setupDatabase dari proyek Apps Script yang terikat pada Google Sheets.'
    );
  }

  properties.setProperty(
    'SMART_EXAM_DATABASE_ID',
    active.getId()
  );

  return active;
}


function seedDefaultData_(ss) {
  const users = ss.getSheetByName('Users');

  if (users.getLastRow() === 1) {
    const salt = Utilities.getUuid();

    users.appendRow([
      createId_(),
      'admin',
      hashPassword_('Admin123!', salt),
      salt,
      'Administrator SMART EXAM',
      '',
      'Super Admin',
      '',
      true,
      now_()
    ]);
  }

  const settings = ss.getSheetByName('SchoolSettings');

  if (settings.getLastRow() === 1) {
    const defaults = [
      ['schoolName', 'SMART EXAM SCHOOL'],
      ['schoolAddress', 'Alamat Sekolah'],
      ['schoolPhone', ''],
      ['schoolEmail', ''],
      ['schoolLogo', ''],
      ['academicYear', '2026/2027'],
      ['appName', APP.NAME]
    ];

    settings.getRange(
      2,
      1,
      defaults.length,
      2
    ).setValues(defaults);
  }
}


/************************************************************
 * API UTAMA
 ************************************************************/

function api(data) {
  try {
    return handlePost_(data || {});
  } catch (error) {
    return {
      status: 'error',
      message: error.message || 'Terjadi kesalahan server.'
    };
  }
}


function handlePost_(data) {
  const action = String(data.action || '');

  switch (action) {

    case 'setupDatabase':
      return setupDatabase();

    case 'login':
      return login_(data);

    case 'register':
      return register_(data);

    case 'logout':
      return logout_(data);

    case 'getDashboard':
      return getDashboard_(data);

    case 'getExams':
      return getExams_(data);

    case 'createExam':
      return createExam_(data);

    case 'saveQuestion':
      return saveQuestion_(data);

    case 'getQuestions':
      return getQuestions_(data);

    case 'updateExamStatus':
      return updateExamStatus_(data);

    case 'startStudentExam':
      return startStudentExam_(data);

    case 'submitExam':
      return submitExam_(data);

    case 'saveViolation':
      return saveViolation_(data);

    case 'getResults':
      return getResults_(data);

    case 'getUsers':
      return getUsers_(data);

    case 'getLogs':
      return getLogs_(data);

    case 'getSchoolSettings':
      return getSchoolSettings_(data);

    case 'saveSchoolSettings':
      return saveSchoolSettings_(data);

    default:
      throw new Error(
        'Aksi tidak dikenal: ' + action
      );
  }
}


/************************************************************
 * AUTHENTICATION
 ************************************************************/

function login_(data) {
  const username =
    String(data.username || '').trim();

  const password =
    String(data.password || '');

  if (!username || !password) {
    throw new Error(
      'Username dan password wajib diisi.'
    );
  }

  const user = findOne_(
    'Users',
    'username',
    username
  );

  if (!user || !isActive_(user.active)) {
    throw new Error(
      'Akun tidak ditemukan atau tidak aktif.'
    );
  }

  const hash = hashPassword_(
    password,
    user.salt
  );

  if (hash !== user.passwordHash) {
    throw new Error(
      'Username atau password salah.'
    );
  }

  const token = createToken_();

  const expiresAt = new Date(
    Date.now() +
    APP.SESSION_HOURS * 60 * 60 * 1000
  );

  append_('Sessions', {
    token: token,
    userId: user.id,
    role: user.role,
    expiresAt: expiresAt,
    createdAt: now_()
  });

  logActivity_(
    user,
    'Login',
    'Pengguna berhasil masuk.'
  );

  return {
    status: 'success',
    token: token,
    user: publicUser_(user)
  };
}


function register_(data) {
  const name = String(data.name || '').trim();
  const username = String(data.username || '').trim();
  const email = String(data.email || '').trim();
  const password = String(data.password || '');
  const className = String(data.className || '').trim();

  if (
    !name ||
    !username ||
    !password ||
    !className
  ) {
    throw new Error(
      'Nama, username, kelas, dan password wajib diisi.'
    );
  }

  if (username.length < 4) {
    throw new Error(
      'Username minimal 4 karakter.'
    );
  }

  if (password.length < 8) {
    throw new Error(
      'Password minimal 8 karakter.'
    );
  }

  if (findOne_('Users', 'username', username)) {
    throw new Error(
      'Username sudah digunakan.'
    );
  }

  const salt = Utilities.getUuid();

  const user = {
    id: createId_(),
    username: username,
    passwordHash: hashPassword_(password, salt),
    salt: salt,
    name: name,
    email: email,
    role: 'Siswa',
    className: className,
    active: true,
    createdAt: now_()
  };

  append_('Users', user);

  logActivity_(
    user,
    'Registrasi',
    'Akun siswa baru dibuat.'
  );

  return {
    status: 'success',
    message: 'Registrasi berhasil. Silakan login.'
  };
}


function logout_(data) {
  const token = String(data.token || '');

  const session = findOne_(
    'Sessions',
    'token',
    token
  );

  if (session) {
    deleteByValue_(
      'Sessions',
      'token',
      token
    );
  }

  return {
    status: 'success',
    message: 'Logout berhasil.'
  };
}


function authenticate_(token) {
  if (!token) {
    throw new Error(
      'Sesi login tidak ditemukan.'
    );
  }

  const session = findOne_(
    'Sessions',
    'token',
    String(token)
  );

  if (!session) {
    throw new Error(
      'Sesi tidak valid. Silakan login kembali.'
    );
  }

  if (
    new Date(session.expiresAt).getTime() <
    Date.now()
  ) {
    deleteByValue_(
      'Sessions',
      'token',
      String(token)
    );

    throw new Error(
      'Sesi telah berakhir. Silakan login kembali.'
    );
  }

  const user = findOne_(
    'Users',
    'id',
    session.userId
  );

  if (!user || !isActive_(user.active)) {
    throw new Error(
      'Akun tidak aktif atau tidak ditemukan.'
    );
  }

  return user;
}


function requireRole_(user, roles) {
  if (roles.indexOf(user.role) === -1) {
    throw new Error(
      'Anda tidak memiliki izin untuk melakukan tindakan ini.'
    );
  }
}


function publicUser_(user) {
  return {
    id: user.id,
    username: user.username,
    name: user.name,
    email: user.email,
    role: user.role,
    className: user.className,
    active: isActive_(user.active)
  };
}


/************************************************************
 * DASHBOARD
 ************************************************************/

function getDashboard_(data) {
  const user = authenticate_(data.token);

  const exams = readAll_('Exams');
  const questions = readAll_('Questions');
  const users = readAll_('Users');
  const results = readAll_('Results');

  let visibleExams = exams;
  let visibleResults = results;

  if (user.role === 'Guru') {
    visibleExams = exams.filter(function(e) {
      return e.teacherId === user.id;
    });

    visibleResults = results.filter(function(r) {
      return visibleExams.some(function(e) {
        return e.id === r.examId;
      });
    });
  }

  if (user.role === 'Siswa') {
    visibleResults = results.filter(function(r) {
      return r.studentId === user.id;
    });
  }

  return {
    status: 'success',
    data: {
      exams: visibleExams.length,
      questions: questions.length,

      students: users.filter(function(u) {
        return u.role === 'Siswa';
      }).length,

      teachers: users.filter(function(u) {
        return u.role === 'Guru';
      }).length,

      activeExams: visibleExams.filter(function(e) {
        return e.status === 'active';
      }).length,

      results: visibleResults.length
    }
  };
}


/************************************************************
 * EXAMS
 ************************************************************/

function getExams_(data) {
  const user = authenticate_(data.token);

  let exams = readAll_('Exams');

  if (user.role === 'Guru') {
    exams = exams.filter(function(e) {
      return e.teacherId === user.id;
    });
  }

  if (user.role === 'Siswa') {
    exams = exams.filter(function(e) {
      return e.status === 'active' ||
             e.status === 'published';
    });
  }

  exams.sort(function(a, b) {
    return new Date(b.createdAt) -
           new Date(a.createdAt);
  });

  return {
    status: 'success',
    data: exams.map(function(e) {
      return {
        id: e.id,
        title: e.title,
        subject: e.subject,
        className: e.className,
        description: e.description,
        duration: Number(e.duration),
        token: user.role === 'Siswa' ? '' : e.token,
        startAt: e.startAt,
        endAt: e.endAt,
        shuffleQuestions: isActive_(e.shuffleQuestions),
        shuffleAnswers: isActive_(e.shuffleAnswers),
        status: e.status,
        teacherId: e.teacherId,
        createdAt: e.createdAt
      };
    })
  };
}


function createExam_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, [
    'Super Admin',
    'Guru'
  ]);

  if (!data.title || !data.subject) {
    throw new Error(
      'Judul dan mata pelajaran wajib diisi.'
    );
  }

  const exam = {
    id: createId_(),
    title: String(data.title).trim(),
    subject: String(data.subject).trim(),
    className: String(data.className || '').trim(),
    description: String(data.description || ''),
    duration: Math.max(
      1,
      Number(data.duration || 60)
    ),
    token: String(
      data.examToken || generateExamToken_()
    ).trim(),
    startAt: data.startAt || '',
    endAt: data.endAt || '',
    shuffleQuestions: Boolean(data.shuffleQuestions),
    shuffleAnswers: Boolean(data.shuffleAnswers),
    status: 'draft',
    teacherId: user.id,
    createdAt: now_()
  };

  append_('Exams', exam);

  logActivity_(
    user,
    'Membuat Ujian',
    exam.title
  );

  return {
    status: 'success',
    data: exam
  };
}


function updateExamStatus_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, [
    'Super Admin',
    'Guru'
  ]);

  const exam = findOne_(
    'Exams',
    'id',
    data.examId
  );

  if (!exam) {
    throw new Error(
      'Ujian tidak ditemukan.'
    );
  }

  if (
    user.role === 'Guru' &&
    exam.teacherId !== user.id
  ) {
    throw new Error(
      'Anda tidak dapat mengubah ujian milik guru lain.'
    );
  }

  const allowed = [
    'draft',
    'active',
    'published',
    'completed',
    'inactive'
  ];

  const status = String(data.status || '');

  if (allowed.indexOf(status) === -1) {
    throw new Error(
      'Status ujian tidak valid.'
    );
  }

  updateById_(
    'Exams',
    exam.id,
    { status: status }
  );

  logActivity_(
    user,
    'Mengubah Status Ujian',
    exam.title + ' → ' + status
  );

  return {
    status: 'success',
    message: 'Status ujian berhasil diperbarui.'
  };
}


/************************************************************
 * QUESTION BANK
 ************************************************************/

function saveQuestion_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, [
    'Super Admin',
    'Guru'
  ]);

  const exam = findOne_(
    'Exams',
    'id',
    data.examId
  );

  if (!exam) {
    throw new Error(
      'Ujian tidak ditemukan.'
    );
  }

  if (
    user.role === 'Guru' &&
    exam.teacherId !== user.id
  ) {
    throw new Error(
      'Anda tidak dapat mengubah soal ujian milik guru lain.'
    );
  }

  const question = String(
    data.question || ''
  ).trim();

  if (!question) {
    throw new Error(
      'Isi pertanyaan wajib diisi.'
    );
  }

  const type = String(
    data.type || 'multiple'
  );

  const options = Array.isArray(data.options)
    ? data.options
    : [];

  if (
    type === 'multiple' &&
    options.length < 2
  ) {
    throw new Error(
      'Soal pilihan ganda minimal memiliki dua pilihan.'
    );
  }

  const answerKey = data.answerKey === undefined
    ? ''
    : String(data.answerKey);

  const id = createId_();

  append_('Questions', {
    id: id,
    examId: String(data.examId),
    question: question,
    type: type,
    options: JSON.stringify(options),
    answerKey: answerKey,
    weight: Math.max(0, Number(data.weight || 1)),
    createdAt: now_()
  });

  logActivity_(
    user,
    'Menambahkan Soal',
    exam.title
  );

  return {
    status: 'success',
    data: {
      id: id
    }
  };
}


function getQuestions_(data) {
  const user = authenticate_(data.token);

  const exam = findOne_(
    'Exams',
    'id',
    data.examId
  );

  if (!exam) {
    throw new Error(
      'Ujian tidak ditemukan.'
    );
  }

  if (
    user.role === 'Guru' &&
    exam.teacherId !== user.id
  ) {
    throw new Error(
      'Akses ditolak.'
    );
  }

  if (
    user.role === 'Siswa' &&
    exam.status !== 'active' &&
    exam.status !== 'published'
  ) {
    throw new Error(
      'Ujian belum tersedia.'
    );
  }

  let questions = readAll_('Questions')
    .filter(function(q) {
      return q.examId === exam.id;
    });

  if (
    user.role === 'Siswa' &&
    isActive_(exam.shuffleQuestions)
  ) {
    questions = shuffleArray_(questions);
  }

  return {
    status: 'success',
    data: questions.map(function(q) {
      return {
        id: q.id,
        examId: q.examId,
        question: q.question,
        type: q.type,
        options: parseJson_(q.options, []),
        weight: Number(q.weight || 1)
      };
    })
  };
}


/************************************************************
 * START STUDENT EXAM
 ************************************************************/

function startStudentExam_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, ['Siswa']);

  const exam = findOne_(
    'Exams',
    'id',
    data.examId
  );

  if (!exam) {
    throw new Error(
      'Ujian tidak ditemukan.'
    );
  }

  if (
    exam.status !== 'active' &&
    exam.status !== 'published'
  ) {
    throw new Error(
      'Ujian belum dibuka.'
    );
  }

  if (
    exam.className &&
    exam.className !== user.className
  ) {
    throw new Error(
      'Ujian ini bukan untuk kelas Anda.'
    );
  }

  const suppliedToken = String(
    data.examToken || ''
  );

  if (
    String(exam.token) !== suppliedToken
  ) {
    throw new Error(
      'Token ujian salah.'
    );
  }

  const current = Date.now();

  if (
    exam.startAt &&
    current < new Date(exam.startAt).getTime()
  ) {
    throw new Error(
      'Ujian belum memasuki jadwal.'
    );
  }

  if (
    exam.endAt &&
    current > new Date(exam.endAt).getTime()
  ) {
    throw new Error(
      'Waktu ujian sudah berakhir.'
    );
  }

  const existingResult = readAll_('Results')
    .find(function(r) {
      return r.examId === exam.id &&
             r.studentId === user.id;
    });

  if (existingResult) {
    throw new Error(
      'Anda sudah mengumpulkan ujian ini.'
    );
  }

  let attempt = readAll_('ExamAttempts')
    .find(function(a) {
      return a.examId === exam.id &&
             a.studentId === user.id &&
             a.status === 'started';
    });

  if (!attempt) {
    attempt = {
      id: createId_(),
      examId: exam.id,
      studentId: user.id,
      startedAt: now_(),
      expiresAt: new Date(
        current +
        Number(exam.duration) * 60000
      ),
      status: 'started',
      submittedAt: ''
    };

    append_('ExamAttempts', attempt);
  }

  if (
    current >
    new Date(attempt.expiresAt).getTime()
  ) {
    throw new Error(
      'Waktu ujian Anda telah habis.'
    );
  }

  let questions = readAll_('Questions')
    .filter(function(q) {
      return q.examId === exam.id;
    });

  if (isActive_(exam.shuffleQuestions)) {
    questions = shuffleArray_(questions);
  }

  const safeQuestions = questions.map(function(q) {
    let options = parseJson_(q.options, []);

    if (isActive_(exam.shuffleAnswers)) {
      options = shuffleArray_(options);
    }

    return {
      id: q.id,
      question: q.question,
      type: q.type,
      options: options,
      weight: Number(q.weight || 1)
    };
  });

  logActivity_(
    user,
    'Memulai Ujian',
    exam.title
  );

  return {
    status: 'success',

    exam: {
      id: exam.id,
      title: exam.title,
      subject: exam.subject,
      duration: Number(exam.duration)
    },

    questions: safeQuestions
  };
}


/************************************************************
 * SUBMIT EXAM
 ************************************************************/

function submitExam_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, ['Siswa']);

  const exam = findOne_(
    'Exams',
    'id',
    data.examId
  );

  if (!exam) {
    throw new Error(
      'Ujian tidak ditemukan.'
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const previous = readAll_('Results')
      .find(function(r) {
        return r.examId === exam.id &&
               r.studentId === user.id;
      });

    if (previous) {
      throw new Error(
        'Jawaban ujian sudah dikumpulkan.'
      );
    }

    const attempt = readAll_('ExamAttempts')
      .find(function(a) {
        return a.examId === exam.id &&
               a.studentId === user.id &&
               a.status === 'started';
      });

    if (!attempt) {
      throw new Error(
        'Sesi ujian tidak ditemukan.'
      );
    }

    const answers = normalizeAnswers_(
      data.answers
    );

    const questions = readAll_('Questions')
      .filter(function(q) {
        return q.examId === exam.id;
      });

    let correct = 0;
    let wrong = 0;
    let totalWeight = 0;
    let earnedWeight = 0;

    const resultId = createId_();

    questions.forEach(function(q) {
      const weight = Number(q.weight || 1);

      totalWeight += weight;

      const answer = answers[q.id];

      let isCorrect = false;
      let score = 0;

      if (q.type === 'multiple') {
        isCorrect =
          String(answer || '').trim().toLowerCase() ===
          String(q.answerKey || '').trim().toLowerCase();

        if (isCorrect) {
          correct++;
          score = weight;
          earnedWeight += weight;
        } else {
          wrong++;
        }
      } else {
        // Esai tidak dinilai otomatis.
        // Penilaian manual dapat dilakukan oleh guru.
        wrong++;
      }

      append_('Answers', {
        id: createId_(),
        resultId: resultId,
        examId: exam.id,
        studentId: user.id,
        questionId: q.id,
        answer: answer === undefined
          ? ''
          : String(answer),
        isCorrect: isCorrect,
        score: score,
        createdAt: now_()
      });
    });

    const score = totalWeight > 0
      ? Math.round(
          earnedWeight / totalWeight * 100
        )
      : 0;

    const result = {
      id: resultId,
      examId: exam.id,
      studentId: user.id,
      studentName: user.name,
      className: user.className,
      score: score,
      correct: correct,
      wrong: wrong,
      total: questions.length,
      submittedAt: now_()
    };

    append_('Results', result);

    updateById_(
      'ExamAttempts',
      attempt.id,
      {
        status: 'submitted',
        submittedAt: now_()
      }
    );

    logActivity_(
      user,
      'Mengumpulkan Ujian',
      exam.title
    );

    return {
      status: 'success',
      result: {
        score: score,
        correct: correct,
        wrong: wrong,
        total: questions.length
      }
    };

  } finally {
    lock.releaseLock();
  }
}


function normalizeAnswers_(answers) {
  if (Array.isArray(answers)) {
    const result = {};

    answers.forEach(function(item) {
      if (item && item.questionId) {
        result[item.questionId] = item.answer;
      }
    });

    return result;
  }

  return answers && typeof answers === 'object'
    ? answers
    : {};
}


/************************************************************
 * VIOLATIONS
 ************************************************************/

function saveViolation_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, ['Siswa']);

  append_('Violations', {
    id: createId_(),
    examId: String(data.examId || ''),
    studentId: user.id,
    type: String(data.type || 'warning'),
    details: String(data.details || ''),
    createdAt: now_()
  });

  logActivity_(
    user,
    'Pelanggaran Ujian',
    String(data.type || 'warning')
  );

  return {
    status: 'success'
  };
}


/************************************************************
 * RESULTS
 ************************************************************/

function getResults_(data) {
  const user = authenticate_(data.token);

  let results = readAll_('Results');

  if (user.role === 'Siswa') {
    results = results.filter(function(r) {
      return r.studentId === user.id;
    });
  }

  if (user.role === 'Guru') {
    const exams = readAll_('Exams').filter(function(e) {
      return e.teacherId === user.id;
    });

    const examIds = exams.map(function(e) {
      return e.id;
    });

    results = results.filter(function(r) {
      return examIds.indexOf(r.examId) !== -1;
    });
  }

  results.sort(function(a, b) {
    return new Date(b.submittedAt) -
           new Date(a.submittedAt);
  });

  return {
    status: 'success',
    data: results
  };
}


/************************************************************
 * USER MANAGEMENT
 ************************************************************/

function getUsers_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, [
    'Super Admin',
    'Guru'
  ]);

  let users = readAll_('Users');

  if (user.role === 'Guru') {
    users = users.filter(function(u) {
      return u.role === 'Siswa' &&
             u.className === user.className;
    });
  }

  return {
    status: 'success',
    data: users.map(publicUser_)
  };
}


/************************************************************
 * ACTIVITY LOGS
 ************************************************************/

function getLogs_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, ['Super Admin']);

  const logs = readAll_('ActivityLogs');

  logs.sort(function(a, b) {
    return new Date(b.createdAt) -
           new Date(a.createdAt);
  });

  return {
    status: 'success',
    data: logs.slice(0, 500)
  };
}


function logActivity_(user, activity, details) {
  append_('ActivityLogs', {
    id: createId_(),
    userId: user.id,
    username: user.username,
    activity: activity,
    details: details,
    createdAt: now_()
  });
}


/************************************************************
 * SCHOOL SETTINGS
 ************************************************************/

function getSchoolSettings_(data) {
  authenticate_(data.token);

  const rows = readAll_('SchoolSettings');

  const settings = {};

  rows.forEach(function(row) {
    settings[row.key] = row.value;
  });

  return {
    status: 'success',
    data: settings
  };
}


function saveSchoolSettings_(data) {
  const user = authenticate_(data.token);

  requireRole_(user, ['Super Admin']);

  const settings = data.settings || {};

  Object.keys(settings).forEach(function(key) {
    const value = settings[key];

    const existing = findOne_(
      'SchoolSettings',
      'key',
      key
    );

    if (existing) {
      updateByValue_(
        'SchoolSettings',
        'key',
        key,
        { value: value }
      );
    } else {
      append_('SchoolSettings', {
        key: key,
        value: value
      });
    }
  });

  logActivity_(
    user,
    'Mengubah Pengaturan Sekolah',
    'Pengaturan sekolah diperbarui.'
  );

  return {
    status: 'success',
    message: 'Pengaturan berhasil disimpan.'
  };
}


/************************************************************
 * SHEET HELPERS
 ************************************************************/

function readAll_(sheetName) {
  const sheet = getDatabase_()
    .getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getValues();

  const headers = values[0];

  return values.slice(1).map(function(row) {
    const obj = {};

    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });

    return obj;
  });
}


function append_(sheetName, object) {
  const sheet = getDatabase_()
    .getSheetByName(sheetName);

  const headers = HEADERS[sheetName];

  if (!headers) {
    throw new Error(
      'Struktur sheet tidak ditemukan: ' + sheetName
    );
  }

  const row = headers.map(function(header) {
    const value = object[header];

    return value === undefined
      ? ''
      : value;
  });

  sheet.appendRow(row);
}


function findOne_(sheetName, field, value) {
  const rows = readAll_(sheetName);

  for (let i = 0; i < rows.length; i++) {
    if (
      String(rows[i][field]) === String(value)
    ) {
      return rows[i];
    }
  }

  return null;
}


function updateById_(sheetName, id, changes) {
  return updateByValue_(
    sheetName,
    'id',
    id,
    changes
  );
}


function updateByValue_(sheetName, field, value, changes) {
  const sheet = getDatabase_()
    .getSheetByName(sheetName);

  const data = sheet.getDataRange().getValues();

  const headers = data[0];

  const fieldIndex = headers.indexOf(field);

  if (fieldIndex === -1) {
    throw new Error(
      'Kolom tidak ditemukan: ' + field
    );
  }

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][fieldIndex]) === String(value)
    ) {
      Object.keys(changes).forEach(function(key) {
        const index = headers.indexOf(key);

        if (index !== -1) {
          sheet.getRange(
            i + 1,
            index + 1
          ).setValue(changes[key]);
        }
      });

      return true;
    }
  }

  return false;
}


function deleteByValue_(sheetName, field, value) {
  const sheet = getDatabase_()
    .getSheetByName(sheetName);

  const data = sheet.getDataRange().getValues();

  const headers = data[0];

  const index = headers.indexOf(field);

  if (index === -1) {
    return false;
  }

  for (let i = data.length - 1; i >= 1; i--) {
    if (
      String(data[i][index]) === String(value)
    ) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }

  return false;
}


/************************************************************
 * SECURITY & UTILITIES
 ************************************************************/

function hashPassword_(password, salt) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + String(password),
    Utilities.Charset.UTF_8
  );

  return digest.map(function(byte) {
    const value = (byte + 256) % 256;

    return ('0' + value.toString(16))
      .slice(-2);
  }).join('');
}


function createId_() {
  return Utilities.getUuid();
}


function createToken_() {
  return Utilities.getUuid() +
         Utilities.getUuid();
}


function generateExamToken_() {
  return Math.random()
    .toString(36)
    .substring(2, 8)
    .toUpperCase();
}


function now_() {
  return Utilities.formatDate(
    new Date(),
    APP.TIMEZONE,
    'yyyy-MM-dd HH:mm:ss'
  );
}


function isActive_(value) {
  return value === true ||
         String(value).toLowerCase() === 'true' ||
         String(value) === '1';
}


function parseJson_(value, fallback) {
  try {
    if (Array.isArray(value)) {
      return value;
    }

    return JSON.parse(value || 'null') || fallback;

  } catch (error) {
    return fallback;
  }
}


function shuffleArray_(array) {
  const result = array.slice();

  for (let i = result.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));

    const temp = result[i];

    result[i] = result[j];
    result[j] = temp;
  }

  return result;
}


/************************************************************
 * TEST DATABASE
 ************************************************************/

function testDatabase() {
  const result = setupDatabase();

  Logger.log(result);

  return result;
}


/************************************************************
 * TEST ADMIN
 ************************************************************/

function testAdmin() {
  const user = findOne_(
    'Users',
    'username',
    'admin'
  );

  Logger.log({
    username: user ? user.username : 'Tidak ditemukan',
    role: user ? user.role : '',
    active: user ? user.active : false
  });
}
