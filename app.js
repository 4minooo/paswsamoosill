const APP_TITLE = "나는 사회복지학전공 근로생이다";
const STORAGE_KEY = "pknu-social-work-office-app:v1";
const USER_KEY = "pknu-social-work-office-users:v1";
const SESSION_KEY = "pknu-social-work-office-session:v1";
const WORKSPACE_ID = "social-welfare-office";

const DAYS = ["월", "화", "수", "목", "금"];
const DAY_INDEX = { 월: 1, 화: 2, 수: 3, 목: 4, 금: 5 };
const VIEW_LABELS = {
  dashboard: "홈",
  shifts: "근로 시간표",
  courses: "수업 세팅",
  calendar: "캘린더",
  todos: "To Do",
  memos: "메모장",
  members: "근로생 명단"
};

const ADMIN_STUDENT_IDS = Array.isArray(window.officeAppConfig?.adminStudentIds)
  ? window.officeAppConfig.adminStudentIds.map(normalizeStudentId)
  : [];

const WORKER_COLORS = [
  "#24756b",
  "#315f9f",
  "#9b6518",
  "#a34055",
  "#6f5d9b",
  "#3f7a4e",
  "#7b4f38",
  "#55748a"
];

const DEFAULT_COURSES = [
  {
    id: "mon-admin-law-no",
    day: "월",
    title: "행정론(노법래)",
    period: "5~7교시",
    time: "13:00~16:00",
    room: "640A",
    setupPerson: "허지은",
    recoverPerson: "허지은",
    grade: "4학년",
    note: ""
  },
  {
    id: "mon-social-security-kim",
    day: "월",
    title: "사회보장론(김선희)",
    period: "6~8교시",
    time: "14:00~17:00",
    room: "C23-304호",
    setupPerson: "김란지",
    recoverPerson: "김란지",
    grade: "3학년",
    note: ""
  },
  {
    id: "mon-case-choi",
    day: "월",
    title: "사례관리론(최수진)",
    period: "글정대",
    time: "시간 확인 필요",
    room: "C23-210호",
    setupPerson: "서민호",
    recoverPerson: "김란지",
    grade: "",
    note: "회수: 화요일"
  },
  {
    id: "tue-research-no",
    day: "화",
    title: "조사론(노법래)",
    period: "2~4교시",
    time: "10:00~13:00",
    room: "634호",
    setupPerson: "김란지",
    recoverPerson: "김수연",
    grade: "2학년",
    note: ""
  },
  {
    id: "tue-human-ha",
    day: "화",
    title: "인행사(허원빈)",
    period: "4~6교시",
    time: "12:00~15:00",
    room: "728호",
    setupPerson: "",
    recoverPerson: "",
    grade: "1학년",
    note: "담당자 확인 필요"
  },
  {
    id: "tue-case-ha",
    day: "화",
    title: "사례관리론(허원빈)",
    period: "7~9교시",
    time: "15:00~18:00",
    room: "635호",
    setupPerson: "김수연",
    recoverPerson: "김수연",
    grade: "3학년",
    note: ""
  },
  {
    id: "tue-intro-kim",
    day: "화",
    title: "사복개론(김은정)",
    period: "9교시",
    time: "17:00~18:00",
    room: "A15-309호",
    setupPerson: "",
    recoverPerson: "",
    grade: "",
    note: "담당자 확인 필요"
  },
  {
    id: "tue-social-security-no",
    day: "화",
    title: "사회보장론(노법래)",
    period: "일반대",
    time: "시간 확인 필요",
    room: "C23-218호",
    setupPerson: "서민호",
    recoverPerson: "이다솜",
    grade: "",
    note: "회수: 수요일"
  },
  {
    id: "tue-admin-kim",
    day: "화",
    title: "행정론(김선희)",
    period: "글정대",
    time: "시간 확인 필요",
    room: "C23-210호",
    setupPerson: "서민호",
    recoverPerson: "이다솜",
    grade: "",
    note: "회수: 수요일"
  },
  {
    id: "wed-welfare-state-kim",
    day: "수",
    title: "복지국가론(김선희)",
    period: "4~6교시",
    time: "12:00~15:00",
    room: "640A",
    setupPerson: "이다솜",
    recoverPerson: "허지은",
    grade: "4학년",
    note: ""
  },
  {
    id: "wed-program-oh",
    day: "수",
    title: "프개평(오영삼)",
    period: "6교시",
    time: "14:00~15:00",
    room: "635호",
    setupPerson: "허지은",
    recoverPerson: "허지은",
    grade: "3학년",
    note: ""
  },
  {
    id: "wed-elderly-oh",
    day: "수",
    title: "노인복지론(오영삼)",
    period: "7~9교시",
    time: "15:00~18:00",
    room: "D15-304호",
    setupPerson: "허지은",
    recoverPerson: "김수연",
    grade: "2학년",
    note: ""
  },
  {
    id: "wed-disability-jeong",
    day: "수",
    title: "장애인복지론(정하은)",
    period: "7~9교시",
    time: "15:00~18:00",
    room: "C23-304호",
    setupPerson: "허지은",
    recoverPerson: "박자민",
    grade: "3학년",
    note: ""
  },
  {
    id: "wed-practice-jeong",
    day: "수",
    title: "실천론(정하은)",
    period: "일반대",
    time: "시간 확인 필요",
    room: "C23-218호",
    setupPerson: "서민호",
    recoverPerson: "김란지",
    grade: "",
    note: "회수: 목요일"
  },
  {
    id: "wed-social-problem-ha",
    day: "수",
    title: "사회문제론(허원빈)",
    period: "글정대",
    time: "시간 확인 필요",
    room: "613호",
    setupPerson: "서민호",
    recoverPerson: "김란지",
    grade: "",
    note: "회수: 목요일"
  },
  {
    id: "thu-photo-ha-oh",
    day: "목",
    title: "포토보이스(허원빈, 오영삼)",
    period: "3~4교시",
    time: "11:00~13:00",
    room: "D15-305호",
    setupPerson: "김란지",
    recoverPerson: "김수연",
    grade: "1학년",
    note: ""
  },
  {
    id: "thu-practice-ha",
    day: "목",
    title: "실천론(허원빈)",
    period: "5~7교시",
    time: "13:00~16:00",
    room: "640A",
    setupPerson: "김란지",
    recoverPerson: "서민호",
    grade: "2학년",
    note: ""
  },
  {
    id: "thu-intro-kim",
    day: "목",
    title: "사복개론(김은정)",
    period: "8~9교시",
    time: "16:00~18:00",
    room: "A15-309호",
    setupPerson: "",
    recoverPerson: "",
    grade: "",
    note: "담당자 확인 필요"
  },
  {
    id: "thu-program-oh",
    day: "목",
    title: "프개평(오영삼)",
    period: "8~9교시",
    time: "16:00~18:00",
    room: "635호",
    setupPerson: "서민호",
    recoverPerson: "서민호",
    grade: "3학년",
    note: ""
  },
  {
    id: "thu-research",
    day: "목",
    title: "사회복지조사론",
    period: "일반대",
    time: "시간 확인 필요",
    room: "C23-218호",
    setupPerson: "서민호",
    recoverPerson: "김란지",
    grade: "",
    note: "회수: 금요일"
  },
  {
    id: "fri-practice-oh",
    day: "금",
    title: "실무역량(오영삼)",
    period: "6~9교시",
    time: "14:00~18:00",
    room: "634호",
    setupPerson: "이다솜",
    recoverPerson: "이다솜",
    grade: "2학년",
    note: ""
  }
];

const DEFAULT_MEMBERS = [];
const DEFAULT_SHIFTS = [];
const LEGACY_SHIFT_IDS = new Set([
  "shift-mon-heo",
  "shift-mon-kim",
  "shift-tue-seo",
  "shift-tue-lee",
  "shift-wed-heo",
  "shift-wed-kim",
  "shift-thu-ranji",
  "shift-thu-seo",
  "shift-fri-lee",
  "shift-fri-park"
]);

let state = createDefaultState();
let currentUser = null;
let activeView = "dashboard";
let authMode = "login";
let selectedCourseId = null;
let calendarCursor = new Date();
let services = { mode: "local", label: "로컬 저장" };
let firebaseApi = null;
let saveTimer = null;

const app = document.querySelector("#app");

boot();

async function boot() {
  services = await initFirebase();
  currentUser = await restoreSession();
  state = await loadState();
  if (currentUser) {
    await reconcileCurrentMember();
  }
  attachGlobalListeners();
  render();
}

async function initFirebase() {
  if (new URL(window.location.href).searchParams.get("local") === "1") {
    return { mode: "local", label: "로컬 저장" };
  }

  const config = window.firebaseConfig;
  if (!config || !config.apiKey || config.apiKey.includes("YOUR_")) {
    return { mode: "local", label: "로컬 저장" };
  }

  try {
    const [
      appModule,
      authModule,
      firestoreModule
    ] = await Promise.all([
      import("https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js"),
      import("https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js"),
      import("https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js")
    ]);

    const firebaseApp = appModule.initializeApp(config);
    firebaseApi = {
      auth: authModule.getAuth(firebaseApp),
      db: firestoreModule.getFirestore(firebaseApp),
      createUserWithEmailAndPassword: authModule.createUserWithEmailAndPassword,
      signInWithEmailAndPassword: authModule.signInWithEmailAndPassword,
      signOut: authModule.signOut,
      onAuthStateChanged: authModule.onAuthStateChanged,
      updateProfile: authModule.updateProfile,
      doc: firestoreModule.doc,
      getDoc: firestoreModule.getDoc,
      setDoc: firestoreModule.setDoc,
      serverTimestamp: firestoreModule.serverTimestamp
    };

    return { mode: "firebase", label: "Firebase 연결" };
  } catch (error) {
    console.warn("Firebase 초기화 실패", error);
    return { mode: "local", label: "로컬 저장" };
  }
}

async function restoreSession() {
  if (services.mode === "firebase" && firebaseApi?.auth) {
    const user = await new Promise((resolve) => {
      const unsubscribe = firebaseApi.onAuthStateChanged(
        firebaseApi.auth,
        (firebaseUser) => {
          unsubscribe();
          resolve(firebaseUser);
        },
        () => {
          unsubscribe();
          resolve(null);
        }
      );
    });
    if (!user) return null;
    const studentId = normalizeStudentId(user.email?.split("@")[0] || user.uid);
    let name = user.displayName || "근로생";
    try {
      const userRef = firebaseApi.doc(firebaseApi.db, "users", studentId);
      const userSnap = await firebaseApi.getDoc(userRef);
      name = userSnap.data()?.name || name;
    } catch (error) {
      console.warn("사용자 프로필 로드 실패", error);
    }
    return {
      id: user.uid,
      name,
      studentId
    };
  }

  const session = readJson(SESSION_KEY, null);
  if (!session?.studentId) {
    return null;
  }

  const users = readJson(USER_KEY, []);
  return users.find((user) => user.studentId === session.studentId) || null;
}

function createDefaultState() {
  return {
    version: 2,
    members: structuredClone(DEFAULT_MEMBERS),
    workers: [],
    shifts: structuredClone(DEFAULT_SHIFTS),
    courses: structuredClone(DEFAULT_COURSES),
    courseStates: {},
    todos: [
      {
        id: "todo-sample-1",
        title: "강의실 포인터 배터리 확인",
        body: "1~4번 포인터 충전 상태 확인 후 부족한 기기는 충전기에 연결",
        dueDate: formatDateInput(new Date()),
        done: false,
        createdAt: nowIso()
      }
    ],
    events: [
      {
        id: "event-sample-1",
        title: "학기 초 강의실 세팅표 재확인",
        date: formatDateInput(new Date()),
        body: "변경된 강의실과 담당자 여부 확인",
        createdAt: nowIso()
      }
    ],
    memos: [
      {
        id: "memo-sample-1",
        text: "공지성 전달사항은 공지로 체크하면 상단에 고정됩니다.",
        notice: true,
        author: "관리",
        createdAt: nowIso()
      }
    ]
  };
}

async function loadState() {
  if (services.mode === "firebase" && firebaseApi && currentUser) {
    try {
      const ref = firebaseApi.doc(firebaseApi.db, "workspaces", WORKSPACE_ID);
      const snap = await firebaseApi.getDoc(ref);
      if (snap.exists()) {
        return mergeState(createDefaultState(), snap.data()?.state || {});
      }
      const fresh = createDefaultState();
      await firebaseApi.setDoc(ref, {
        state: fresh,
        updatedAt: firebaseApi.serverTimestamp()
      });
      return fresh;
    } catch (error) {
      console.warn("Firestore 상태 로드 실패", error);
    }
  }

  return mergeState(createDefaultState(), readJson(STORAGE_KEY, {}));
}

function mergeState(base, saved) {
  const migratedMembers = Array.isArray(saved.members) ? saved.members : [];
  const savedShifts = Array.isArray(saved.shifts) ? saved.shifts : base.shifts;
  const migratedShifts = saved.version === 2
    ? savedShifts
    : savedShifts.filter((shift) => !LEGACY_SHIFT_IDS.has(shift.id));

  return {
    ...base,
    ...saved,
    version: 2,
    members: migratedMembers,
    workers: [],
    shifts: migratedShifts,
    courses: saved.courses?.length ? saved.courses : base.courses,
    courseStates: saved.courseStates || base.courseStates,
    todos: saved.todos || base.todos,
    events: saved.events || base.events,
    memos: saved.memos || base.memos
  };
}

function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    if (services.mode === "firebase" && firebaseApi) {
      try {
        const ref = firebaseApi.doc(firebaseApi.db, "workspaces", WORKSPACE_ID);
        await firebaseApi.setDoc(
          ref,
          {
            state,
            updatedAt: firebaseApi.serverTimestamp()
          },
          { merge: true }
        );
        return;
      } catch (error) {
        console.warn("Firestore 저장 실패", error);
      }
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, 180);
}

function attachGlobalListeners() {
  app.addEventListener("click", async (event) => {
    const target = event.target.closest("[data-action], [data-nav], [data-auth-mode], [data-course-id], [data-month]");
    if (!target) return;

    if (target.dataset.authMode) {
      authMode = target.dataset.authMode;
      render();
      return;
    }

    if (target.dataset.nav) {
      activeView = target.dataset.nav;
      selectedCourseId = null;
      render();
      return;
    }

    if (target.dataset.courseId) {
      selectedCourseId = target.dataset.courseId;
      render();
      return;
    }

    if (target.dataset.month) {
      calendarCursor = new Date(calendarCursor.getFullYear(), calendarCursor.getMonth() + Number(target.dataset.month), 1);
      render();
      return;
    }

    const { action } = target.dataset;
    if (action === "logout") {
      await logout();
    }
    if (action === "close-drawer") {
      selectedCourseId = null;
      render();
    }
    if (action === "recover-course") {
      recoverCourse(target.dataset.id);
    }
    if (action === "reset-course") {
      resetCourse(target.dataset.id);
    }
    if (action === "delete-shift") {
      deleteShift(target.dataset.id);
    }
    if (action === "toggle-todo") {
      toggleTodo(target.dataset.id);
    }
    if (action === "delete-todo") {
      deleteTodo(target.dataset.id);
    }
    if (action === "delete-memo") {
      deleteMemo(target.dataset.id);
    }
    if (action === "delete-event") {
      deleteEvent(target.dataset.id);
    }
    if (action === "approve-member") {
      approveMember(target.dataset.id);
    }
    if (action === "ban-member") {
      banMember(target.dataset.id);
    }
    if (action === "restore-member") {
      restoreMember(target.dataset.id);
    }
  });

  app.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.target;
    const type = form.dataset.form;
    if (type === "auth") {
      await handleAuth(form);
    }
    if (type === "shift") {
      addShift(form);
    }
    if (type === "course-setup") {
      setupCourse(form);
    }
    if (type === "todo") {
      addTodo(form);
    }
    if (type === "memo") {
      addMemo(form);
    }
    if (type === "event") {
      addEvent(form);
    }
  });
}

function render(message = "") {
  if (!currentUser) {
    renderAuth(message);
    return;
  }

  const membership = getCurrentMember();
  if (!membership || membership.status !== "approved") {
    renderAccessGate(membership);
    return;
  }

  const view = renderView();
  app.innerHTML = `
    <div class="app-shell">
      ${renderSidebar()}
      <main class="content">
        ${renderMobileTopbar()}
        <div class="page-header">
          <div>
            <h1>${VIEW_LABELS[activeView]}</h1>
            <p>${renderViewSubtitle()}</p>
          </div>
          <span class="connection-badge">
            <span class="connection-dot ${services.mode === "firebase" ? "live" : ""}"></span>
            ${services.label}
          </span>
        </div>
        ${message ? `<div class="top-message">${escapeHtml(message)}</div>` : ""}
        ${view}
      </main>
    </div>
    ${selectedCourseId ? renderCourseDrawer(selectedCourseId) : ""}
  `;
}

function renderAuth(message = "") {
  const isSignup = authMode === "signup";
  app.innerHTML = `
    <main class="auth-shell">
      <section class="auth-panel">
        <h1 class="auth-title">${APP_TITLE}</h1>
        <p class="auth-subtitle">전공사무실 근로 업무 관리</p>
        <div class="auth-tabs">
          <button class="${!isSignup ? "active" : ""}" type="button" data-auth-mode="login">로그인</button>
          <button class="${isSignup ? "active" : ""}" type="button" data-auth-mode="signup">회원가입</button>
        </div>
        <form data-form="auth">
          <div class="field ${isSignup ? "" : "hidden"}">
            <label for="name">이름</label>
            <input id="name" name="name" autocomplete="name" placeholder="홍길동" />
          </div>
          <div class="field">
            <label for="studentId">학번</label>
            <input id="studentId" name="studentId" autocomplete="username" inputmode="numeric" required placeholder="202600000" />
          </div>
          <div class="field">
            <label for="password">비밀번호</label>
            <input id="password" name="password" type="password" autocomplete="${isSignup ? "new-password" : "current-password"}" required minlength="6" placeholder="6자 이상" />
          </div>
          <button class="primary-button" type="submit">${isSignup ? "가입하기" : "로그인"}</button>
          <p class="status-line">${escapeHtml(message)}</p>
        </form>
      </section>
    </main>
  `;
}

function renderAccessGate(membership) {
  const isBanned = membership?.status === "banned";
  app.innerHTML = `
    <main class="auth-shell">
      <section class="auth-panel">
        <h1 class="auth-title">${APP_TITLE}</h1>
        <p class="auth-subtitle">${escapeHtml(currentUser.name)} · ${escapeHtml(currentUser.studentId)}</p>
        <div class="top-message">
          ${isBanned ? "이 계정은 근로생 명단에서 제외되었습니다." : "회원가입 신청이 접수되었습니다. 관리자의 승인을 기다려 주세요."}
        </div>
        <p class="small-muted">
          ${isBanned ? "다시 접근이 필요하면 관리자에게 복구를 요청하세요." : "승인되면 근로생 명단에 자동으로 추가되고 앱 기능을 사용할 수 있습니다."}
        </p>
        <button class="primary-button" type="button" data-action="logout">로그아웃</button>
      </section>
    </main>
  `;
}

function renderSidebar() {
  return `
    <aside class="sidebar">
      <div class="brand">
        <strong>${APP_TITLE}</strong>
        <span>국립부경대학교 사회복지학전공</span>
      </div>
      <nav class="nav-list">
        ${renderNavButtons()}
      </nav>
      <div class="sidebar-footer">
        <div class="user-pill">
          <strong>${escapeHtml(currentUser.name)}</strong>
          <span>${escapeHtml(currentUser.studentId)}</span>
        </div>
        <button class="ghost-button" type="button" data-action="logout">로그아웃</button>
      </div>
    </aside>
  `;
}

function renderMobileTopbar() {
  return `
    <div class="mobile-topbar">
      <div class="brand">
        <strong>${APP_TITLE}</strong>
        <span>${escapeHtml(currentUser.name)} · ${escapeHtml(currentUser.studentId)}</span>
      </div>
      <nav class="mobile-nav">
        ${renderNavButtons()}
        <button class="nav-button" type="button" data-action="logout">
          <span class="nav-icon">↗</span>
          로그아웃
        </button>
      </nav>
    </div>
  `;
}

function renderNavButtons() {
  const icons = {
    dashboard: "⌂",
    shifts: "▦",
    courses: "✓",
    calendar: "□",
    todos: "!",
    memos: "✎",
    members: "◇"
  };
  return Object.entries(VIEW_LABELS)
    .map(
      ([view, label]) => `
        <button class="nav-button ${activeView === view ? "active" : ""}" type="button" data-nav="${view}">
          <span class="nav-icon">${icons[view]}</span>
          ${label}
        </button>
      `
    )
    .join("");
}

function renderViewSubtitle() {
  const today = new Date();
  const day = getKoreanDay(today);
  const subtitles = {
    dashboard: `${formatKoreanDate(today)} · 오늘 요일 ${day}`,
    shifts: "근로생별 주간 근무 배치",
    courses: "포인터 세팅과 회수 현황",
    calendar: "일정과 마감일",
    todos: "업무 목록",
    memos: "공지와 전달사항",
    members: "승인된 회원이 근로생 명단입니다"
  };
  return subtitles[activeView];
}

function renderView() {
  if (activeView === "dashboard") return renderDashboard();
  if (activeView === "shifts") return renderShiftView();
  if (activeView === "courses") return renderCourseView();
  if (activeView === "calendar") return renderCalendarView();
  if (activeView === "todos") return renderTodoView();
  if (activeView === "memos") return renderMemoView();
  if (activeView === "members") return renderMemberView();
  return renderDashboard();
}

function renderDashboard() {
  const todayDay = getKoreanDay(new Date());
  const todayShifts = getShiftsForDay(todayDay);
  const pendingTodos = state.todos.filter((todo) => !todo.done).sort(sortByDueDate).slice(0, 5);
  const notices = state.memos.filter((memo) => memo.notice).sort(sortByCreatedDesc).slice(0, 4);
  const pendingCourses = state.courses.filter((course) => !isCourseRecovered(course.id)).length;
  const activePointers = getPointerStatuses().filter((pointer) => pointer.course).length;

  return `
    <section class="stat-grid">
      ${renderStat(pendingTodos.length, "남은 업무")}
      ${renderStat(todayShifts.length, "오늘 근로")}
      ${renderStat(pendingCourses, "회수 확인 필요")}
      ${renderStat(activePointers, "사용 중 포인터")}
    </section>
    ${renderPointerStrip()}
    <section class="dashboard-grid">
      <div class="panel">
        <div class="panel-header">
          <h2>오늘 근로 시간표</h2>
          <button class="secondary-button" type="button" data-nav="shifts">전체 보기</button>
        </div>
        <div class="panel-body">
          ${renderShiftList(todayShifts)}
        </div>
      </div>
      <div class="panel">
        <div class="panel-header">
          <h2>To Do</h2>
          <button class="secondary-button" type="button" data-nav="todos">관리</button>
        </div>
        <div class="panel-body">
          ${renderTodoList(pendingTodos)}
        </div>
      </div>
      <div class="panel">
        <div class="panel-header">
          <h2>공지</h2>
          <button class="secondary-button" type="button" data-nav="memos">메모장</button>
        </div>
        <div class="panel-body">
          ${renderMemoList(notices)}
        </div>
      </div>
      <div class="panel">
        <div class="panel-header">
          <h2>오늘 수업 세팅</h2>
          <button class="secondary-button" type="button" data-nav="courses">세팅표</button>
        </div>
        <div class="panel-body">
          ${renderCourseCards(state.courses.filter((course) => course.day === todayDay))}
        </div>
      </div>
    </section>
  `;
}

function renderShiftView() {
  const currentMember = getCurrentMember();
  return `
    <section class="two-column">
      <div class="panel">
        <div class="panel-header">
          <h2>내 근로 일정 추가</h2>
        </div>
        <div class="panel-body">
          <form data-form="shift" class="form-grid">
            <div class="field full">
              <label>근로생</label>
              <input value="${escapeAttr(currentMember?.name || currentUser.name)}" disabled />
            </div>
            <div class="field">
              <label for="shiftDay">요일</label>
              <select id="shiftDay" name="day" required>
                ${DAYS.map((day) => `<option value="${day}">${day}</option>`).join("")}
              </select>
            </div>
            <div class="field">
              <label for="shiftStart">시작 시간</label>
              <select id="shiftStart" name="start" required>
                ${renderTimeOptions("09:00")}
              </select>
            </div>
            <div class="field">
              <label for="shiftEnd">종료 시간</label>
              <select id="shiftEnd" name="end" required>
                ${renderTimeOptions("18:00")}
              </select>
            </div>
            <label class="checkbox-field full">
              <input name="repeat" type="checkbox" checked />
              매주 반복
            </label>
            <div class="full">
              <button class="primary-button" type="submit">추가</button>
            </div>
          </form>
        </div>
      </div>
      <div class="panel">
        <div class="panel-header">
          <h2>근로생 색상</h2>
        </div>
        <div class="panel-body">
          <div class="shift-list">
            ${getApprovedMembers().length ? getApprovedMembers()
              .map(
                (worker) => `
                  <div class="shift-card" style="--worker-color: ${worker.color}">
                    <div class="item-head">
                      <strong>${escapeHtml(worker.name)}</strong>
                      <span class="badge">${escapeHtml(worker.studentId || "등록 대기")}</span>
                    </div>
                  </div>
                `
              )
              .join("") : `<div class="empty-state">승인된 근로생이 없습니다.</div>`}
          </div>
        </div>
      </div>
    </section>
    <section class="panel" style="margin-top:16px;">
      <div class="panel-header">
        <h2>일주일 근로 시간표</h2>
        <span class="badge teal">00~24시 기준</span>
      </div>
      <div class="panel-body">
        ${renderShiftTimetable()}
      </div>
    </section>
  `;
}

function renderShiftTimetable() {
  return `
    <div class="timetable shift-timetable">
      ${DAYS.map((day) => {
        const shifts = getShiftsForDay(day);
        return `
          <div class="day-column shift-day">
            <div class="day-title">${day}<span>${shifts.length}건</span></div>
            ${renderDayTimeline(shifts)}
          </div>
        `;
      }).join("")}
    </div>
  `;
}

function renderDayTimeline(shifts) {
  return `
    <div class="timeline-wrap">
      <div class="timeline-grid">
        ${Array.from({ length: 25 }, (_, hour) => `<span class="hour-label">${String(hour).padStart(2, "0")}:00</span>`).join("")}
      </div>
      <div class="timeline-lane">
        ${shifts.sort(sortByShiftTime).map(renderShiftBlock).join("")}
      </div>
    </div>
  `;
}

function renderShiftBlock(shift) {
  const worker = findWorker(shift.workerName, shift.studentId);
  const start = minutesFromHHMM(shift.start);
  const end = minutesFromHHMM(shift.end);
  const top = Math.max(0, (start / 1440) * 100);
  const height = Math.max(3.2, ((end - start) / 1440) * 100);
  const canDelete = canManageShift(shift);

  return `
    <div class="shift-block" style="--worker-color: ${worker?.color || WORKER_COLORS[0]}; top: ${top}%; height: ${height}%;">
      <div class="item-head">
        <div>
          <strong>${escapeHtml(shift.workerName)}</strong>
          <span>${escapeHtml(shift.start)}~${escapeHtml(shift.end)}</span>
        </div>
        ${canDelete ? `<button class="icon-button mini" type="button" data-action="delete-shift" data-id="${shift.id}" aria-label="삭제">×</button>` : ""}
      </div>
    </div>
  `;
}

function renderCourseView() {
  return `
    ${renderPointerStrip()}
    <section class="panel">
      <div class="panel-header">
        <h2>2026-1학기 수업 세팅표</h2>
        <span class="badge teal">포인터 1~4번</span>
      </div>
      <div class="panel-body">
        ${renderCourseTimetable()}
      </div>
    </section>
  `;
}

function renderCourseTimetable() {
  return `
    <div class="timetable">
      ${DAYS.map((day) => {
        const courses = state.courses.filter((course) => course.day === day).sort(sortByCourseTime);
        return `
          <div class="day-column">
            <div class="day-title">${day}<span>${courses.length}개 수업</span></div>
            <div class="course-list">
              ${renderCourseCards(courses)}
            </div>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

function renderCourseCards(courses) {
  if (!courses.length) {
    return `<div class="empty-state">표시할 수업이 없습니다.</div>`;
  }
  return courses
    .map((course) => {
      const status = getCourseStatus(course.id);
      const stateClass = status.kind === "done" ? "done" : status.kind === "no-pointer" ? "no-pointer" : status.kind === "setup" ? "setup" : "";
      return `
        <button class="course-card ${stateClass}" type="button" data-course-id="${course.id}">
          <span class="course-name">${escapeHtml(course.title)}</span>
          <span class="course-meta">${escapeHtml(course.time)} · ${escapeHtml(course.room)}</span>
          <span class="course-meta">세팅 ${escapeHtml(course.setupPerson || "확인 필요")} · 회수 ${escapeHtml(course.recoverPerson || "확인 필요")}</span>
          <span class="badge-row">
            <span class="badge ${status.badgeClass}">${status.label}</span>
            ${course.grade ? `<span class="badge">${escapeHtml(course.grade)}</span>` : ""}
          </span>
        </button>
      `;
    })
    .join("");
}

function renderPointerStrip() {
  return `
    <section class="pointer-strip">
      ${getPointerStatuses()
        .map(
          (pointer) => `
            <div class="pointer-card ${pointer.course ? "busy" : ""}">
              <strong>포인터 ${pointer.no}</strong>
              <span>${pointer.course ? `${escapeHtml(pointer.course.title)} · ${escapeHtml(pointer.course.room)}` : "대기 중"}</span>
            </div>
          `
        )
        .join("")}
    </section>
  `;
}

function renderCalendarView() {
  const events = state.events.sort(sortByDateAsc);
  return `
    <section class="two-column">
      <div class="panel">
        <div class="panel-header">
          <h2>일정 추가</h2>
        </div>
        <div class="panel-body">
          <form data-form="event" class="form-grid">
            <div class="field full">
              <label for="eventTitle">일정 제목</label>
              <input id="eventTitle" name="title" required placeholder="예: 자료 제출" />
            </div>
            <div class="field">
              <label for="eventDate">날짜</label>
              <input id="eventDate" name="date" type="date" value="${formatDateInput(new Date())}" required />
            </div>
            <div class="field full">
              <label for="eventBody">내용</label>
              <textarea id="eventBody" name="body" placeholder="필요한 내용을 적어두세요."></textarea>
            </div>
            <div class="full">
              <button class="primary-button" type="submit">추가</button>
            </div>
          </form>
        </div>
      </div>
      <div class="panel">
        <div class="panel-header">
          <h2>다가오는 일정</h2>
        </div>
        <div class="panel-body">
          ${renderEventList(events.slice(0, 8))}
        </div>
      </div>
    </section>
    <section class="panel" style="margin-top:16px;">
      <div class="panel-header">
        <div class="calendar-head">
          <button class="icon-button" type="button" data-month="-1" aria-label="이전 달">‹</button>
          <h2>${calendarCursor.getFullYear()}년 ${calendarCursor.getMonth() + 1}월</h2>
          <button class="icon-button" type="button" data-month="1" aria-label="다음 달">›</button>
        </div>
      </div>
      <div class="panel-body">
        <div class="calendar-shell">
          ${renderCalendarGrid()}
        </div>
      </div>
    </section>
  `;
}

function renderCalendarGrid() {
  const year = calendarCursor.getFullYear();
  const month = calendarCursor.getMonth();
  const first = new Date(year, month, 1);
  const start = new Date(year, month, 1 - first.getDay());
  const todayKey = formatDateInput(new Date());
  const weekdays = ["일", "월", "화", "수", "목", "금", "토"];
  const cells = [];

  for (let i = 0; i < 42; i += 1) {
    const date = new Date(start.getFullYear(), start.getMonth(), start.getDate() + i);
    const key = formatDateInput(date);
    const dayEvents = state.events.filter((event) => event.date === key);
    const dayTodos = state.todos.filter((todo) => todo.dueDate === key);
    cells.push(`
      <div class="calendar-day ${date.getMonth() === month ? "" : "outside"} ${key === todayKey ? "today" : ""}">
        <span class="date-num">${date.getDate()}</span>
        ${dayEvents.map((event) => `<span class="calendar-chip">${escapeHtml(event.title)}</span>`).join("")}
        ${dayTodos.map((todo) => `<span class="calendar-chip todo">${escapeHtml(todo.title)}</span>`).join("")}
      </div>
    `);
  }

  return `
    <div class="calendar-grid">
      ${weekdays.map((day) => `<div class="calendar-weekday">${day}</div>`).join("")}
      ${cells.join("")}
    </div>
  `;
}

function renderTodoView() {
  return `
    <section class="two-column">
      <div class="panel">
        <div class="panel-header">
          <h2>해야 할 일 추가</h2>
        </div>
        <div class="panel-body">
          <form data-form="todo" class="form-grid">
            <div class="field full">
              <label for="todoTitle">업무 제목</label>
              <input id="todoTitle" name="title" required placeholder="예: 강의실 세팅 확인" />
            </div>
            <div class="field">
              <label for="todoDue">마감 기한</label>
              <input id="todoDue" name="dueDate" type="date" value="${formatDateInput(new Date())}" required />
            </div>
            <div class="field full">
              <label for="todoBody">업무 내용</label>
              <textarea id="todoBody" name="body" required placeholder="업무 내용을 입력하세요."></textarea>
            </div>
            <div class="full">
              <button class="primary-button" type="submit">추가</button>
            </div>
          </form>
        </div>
      </div>
      <div class="panel">
        <div class="panel-header">
          <h2>업무 목록</h2>
        </div>
        <div class="panel-body">
          ${renderTodoList(state.todos.sort(sortByDueDate))}
        </div>
      </div>
    </section>
  `;
}

function renderMemoView() {
  const memos = state.memos.sort((a, b) => Number(b.notice) - Number(a.notice) || sortByCreatedDesc(a, b));
  return `
    <section class="two-column">
      <div class="panel">
        <div class="panel-header">
          <h2>메모 작성</h2>
        </div>
        <div class="panel-body">
          <form data-form="memo">
            <div class="field">
              <label for="memoText">전달사항</label>
              <textarea id="memoText" name="text" required placeholder="공지 또는 전달사항을 입력하세요."></textarea>
            </div>
            <label class="checkbox-field">
              <input name="notice" type="checkbox" />
              공지로 고정
            </label>
            <button class="primary-button" type="submit">등록</button>
          </form>
        </div>
      </div>
      <div class="panel">
        <div class="panel-header">
          <h2>메모장</h2>
        </div>
        <div class="panel-body">
          ${renderMemoList(memos)}
        </div>
      </div>
    </section>
  `;
}

function renderMemberView() {
  const members = getMembersSorted();
  const approved = members.filter((member) => member.status === "approved");
  const pending = members.filter((member) => member.status === "pending");
  const banned = members.filter((member) => member.status === "banned");
  const admin = isCurrentAdmin();

  return `
    <section class="stat-grid">
      ${renderStat(approved.length, "승인 근로생")}
      ${renderStat(pending.length, "승인 대기")}
      ${renderStat(banned.length, "추방 계정")}
      ${renderStat(approved.filter((member) => member.role === "admin").length, "관리자")}
    </section>
    <section class="two-column">
      <div class="panel">
        <div class="panel-header">
          <h2>승인된 근로생</h2>
        </div>
        <div class="panel-body">
          ${renderMemberList(approved, admin)}
        </div>
      </div>
      <div class="panel">
        <div class="panel-header">
          <h2>승인 대기</h2>
          ${admin ? `<span class="badge amber">관리자 승인 필요</span>` : ""}
        </div>
        <div class="panel-body">
          ${renderMemberList(pending, admin)}
        </div>
      </div>
    </section>
    <section class="panel" style="margin-top:16px;">
      <div class="panel-header">
        <h2>추방된 회원</h2>
      </div>
      <div class="panel-body">
        ${renderMemberList(banned, admin)}
      </div>
    </section>
  `;
}

function renderStat(value, label) {
  return `
    <div class="stat">
      <strong>${value}</strong>
      <span>${label}</span>
    </div>
  `;
}

function renderShiftList(shifts) {
  if (!shifts.length) {
    return `<div class="empty-state">등록된 근로 일정이 없습니다.</div>`;
  }

  return `
    <div class="shift-list">
      ${shifts
        .sort(sortByShiftTime)
        .map((shift) => {
          const worker = findWorker(shift.workerName, shift.studentId);
          const canDelete = canManageShift(shift);
          return `
            <div class="shift-card" style="--worker-color: ${worker?.color || WORKER_COLORS[0]}">
              <div class="item-head">
                <div>
                  <strong>${escapeHtml(shift.workerName)}</strong>
                  <div class="small-muted">${escapeHtml(shift.day)} · ${escapeHtml(shift.start)}~${escapeHtml(shift.end)} ${shift.repeat ? "· 매주" : ""}</div>
                </div>
                <div class="item-actions">
                  ${canDelete ? `<button class="icon-button" type="button" data-action="delete-shift" data-id="${shift.id}" aria-label="삭제">×</button>` : ""}
                </div>
              </div>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderTodoList(todos) {
  if (!todos.length) {
    return `<div class="empty-state">등록된 업무가 없습니다.</div>`;
  }

  return `
    <div class="todo-list">
      ${todos
        .map(
          (todo) => `
            <div class="todo-item ${todo.done ? "done" : ""}">
              <div class="item-head">
                <div>
                  <strong>${escapeHtml(todo.title)}</strong>
                  <div class="small-muted">마감 ${escapeHtml(todo.dueDate)}</div>
                </div>
                <div class="item-actions">
                  <button class="icon-button" type="button" data-action="toggle-todo" data-id="${todo.id}" aria-label="완료">${todo.done ? "↺" : "✓"}</button>
                  <button class="icon-button" type="button" data-action="delete-todo" data-id="${todo.id}" aria-label="삭제">×</button>
                </div>
              </div>
              <p class="memo-text">${escapeHtml(todo.body)}</p>
            </div>
          `
        )
        .join("")}
    </div>
  `;
}

function renderMemoList(memos) {
  if (!memos.length) {
    return `<div class="empty-state">등록된 메모가 없습니다.</div>`;
  }

  return `
    <div class="memo-list">
      ${memos
        .map(
          (memo) => `
            <div class="memo-item ${memo.notice ? "notice" : ""}">
              <div class="item-head">
                <div>
                  <strong>${memo.notice ? "공지" : "메모"}</strong>
                  <div class="small-muted">${escapeHtml(memo.author)} · ${formatKoreanDateTime(memo.createdAt)}</div>
                </div>
                <div class="item-actions">
                  <button class="icon-button" type="button" data-action="delete-memo" data-id="${memo.id}" aria-label="삭제">×</button>
                </div>
              </div>
              <p class="memo-text">${escapeHtml(memo.text)}</p>
            </div>
          `
        )
        .join("")}
    </div>
  `;
}

function renderMemberList(members, admin) {
  if (!members.length) {
    return `<div class="empty-state">표시할 회원이 없습니다.</div>`;
  }

  return `
    <div class="memo-list">
      ${members
        .map((member) => {
          const canBan = admin && member.status === "approved" && member.studentId !== currentUser.studentId;
          const canApprove = admin && member.status === "pending";
          const canRestore = admin && member.status === "banned";
          return `
            <div class="shift-card" style="--worker-color: ${member.color || WORKER_COLORS[0]}">
              <div class="item-head">
                <div>
                  <strong>${escapeHtml(member.name)}</strong>
                  <div class="small-muted">${escapeHtml(member.studentId)} · ${renderMemberStatus(member)}</div>
                </div>
                <div class="item-actions">
                  ${canApprove ? `<button class="secondary-button" type="button" data-action="approve-member" data-id="${member.studentId}">승인</button>` : ""}
                  ${canBan ? `<button class="danger-button" type="button" data-action="ban-member" data-id="${member.studentId}">추방</button>` : ""}
                  ${canRestore ? `<button class="secondary-button" type="button" data-action="restore-member" data-id="${member.studentId}">복구</button>` : ""}
                </div>
              </div>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderMemberStatus(member) {
  const role = member.role === "admin" ? "관리자" : "근로생";
  if (member.status === "pending") return "승인 대기";
  if (member.status === "banned") return "추방됨";
  return role;
}

function renderEventList(events) {
  if (!events.length) {
    return `<div class="empty-state">등록된 일정이 없습니다.</div>`;
  }

  return `
    <div class="event-list">
      ${events
        .map(
          (event) => `
            <div class="event-item">
              <div class="item-head">
                <div>
                  <strong>${escapeHtml(event.title)}</strong>
                  <div class="small-muted">${escapeHtml(event.date)}</div>
                </div>
                <div class="item-actions">
                  <button class="icon-button" type="button" data-action="delete-event" data-id="${event.id}" aria-label="삭제">×</button>
                </div>
              </div>
              ${event.body ? `<p class="memo-text">${escapeHtml(event.body)}</p>` : ""}
            </div>
          `
        )
        .join("")}
    </div>
  `;
}

function renderCourseDrawer(courseId) {
  const course = state.courses.find((item) => item.id === courseId);
  if (!course) return "";
  const record = state.courseStates[course.id];
  const status = getCourseStatus(course.id);
  const canRecover = record?.setupAt && !record?.recoveredAt;

  return `
    <div class="drawer-backdrop">
      <aside class="drawer" role="dialog" aria-modal="true">
        <div class="drawer-header">
          <div>
            <h2>${escapeHtml(course.title)}</h2>
            <span class="badge ${status.badgeClass}">${status.label}</span>
          </div>
          <button class="icon-button" type="button" data-action="close-drawer" aria-label="닫기">×</button>
        </div>
        <div class="drawer-body">
          <div class="detail-list">
            ${renderDetail("요일", course.day)}
            ${renderDetail("교시", course.period)}
            ${renderDetail("시간", course.time)}
            ${renderDetail("강의실", course.room)}
            ${renderDetail("세팅", course.setupPerson || "확인 필요")}
            ${renderDetail("회수", course.recoverPerson || "확인 필요")}
            ${course.grade ? renderDetail("학년", course.grade) : ""}
            ${course.note ? renderDetail("비고", course.note) : ""}
            ${record?.setupAt ? renderDetail("설치 기록", `${record.pointer === "none" ? "포인터 설치 안 함" : `포인터 ${record.pointer}번`} · ${record.setupBy} · ${formatKoreanDateTime(record.setupAt)}`) : ""}
            ${record?.recoveredAt ? renderDetail("회수 기록", `${record.recoveredBy} · ${formatKoreanDateTime(record.recoveredAt)}`) : ""}
          </div>
          ${
            !record?.setupAt
              ? `
                <form data-form="course-setup" class="form-grid">
                  <input type="hidden" name="courseId" value="${course.id}" />
                  <div class="field full">
                    <label for="pointerNo">설치한 포인터</label>
                    <select id="pointerNo" name="pointer" required>
                      <option value="1">포인터 1번</option>
                      <option value="2">포인터 2번</option>
                      <option value="3">포인터 3번</option>
                      <option value="4">포인터 4번</option>
                      <option value="none">포인터 설치 안 함</option>
                    </select>
                  </div>
                  <div class="full">
                    <button class="primary-button" type="submit">세팅 완료</button>
                  </div>
                </form>
              `
              : `
                <div class="action-row">
                  ${
                    canRecover
                      ? `<button class="primary-button" type="button" data-action="recover-course" data-id="${course.id}">회수 완료</button>`
                      : ""
                  }
                  <button class="danger-button" type="button" data-action="reset-course" data-id="${course.id}">기록 초기화</button>
                </div>
              `
          }
        </div>
      </aside>
    </div>
  `;
}

function renderDetail(label, value) {
  return `
    <div class="detail-row">
      <span>${escapeHtml(label)}</span>
      <span>${escapeHtml(value)}</span>
    </div>
  `;
}

async function handleAuth(form) {
  const data = new FormData(form);
  const name = String(data.get("name") || "").trim();
  const studentId = normalizeStudentId(String(data.get("studentId") || ""));
  const password = String(data.get("password") || "");

  if (!studentId || !password) {
    renderAuth("학번과 비밀번호를 입력하세요.");
    return;
  }

  if (authMode === "signup" && !name) {
    renderAuth("이름을 입력하세요.");
    return;
  }

  try {
    if (services.mode === "firebase" && firebaseApi) {
      await handleFirebaseAuth({ name, studentId, password });
    } else {
      handleLocalAuth({ name, studentId, password });
    }
    await reconcileCurrentMember({ signup: authMode === "signup" });
    scheduleSave();
    render();
  } catch (error) {
    renderAuth(error.message || "로그인 처리 중 오류가 발생했습니다.");
  }
}

async function handleFirebaseAuth({ name, studentId, password }) {
  const email = `${studentId}@pknu-work.app`;

  if (authMode === "signup") {
    const credential = await firebaseApi.createUserWithEmailAndPassword(firebaseApi.auth, email, password);
    await firebaseApi.updateProfile(credential.user, { displayName: name });
    currentUser = { id: credential.user.uid, name, studentId };
    const userRef = firebaseApi.doc(firebaseApi.db, "users", studentId);
    await firebaseApi.setDoc(
      userRef,
      {
        name,
        studentId,
        updatedAt: firebaseApi.serverTimestamp()
      },
      { merge: true }
    );
  } else {
    const credential = await firebaseApi.signInWithEmailAndPassword(firebaseApi.auth, email, password);
    const userRef = firebaseApi.doc(firebaseApi.db, "users", studentId);
    const userSnap = await firebaseApi.getDoc(userRef);
    currentUser = {
      id: credential.user.uid,
      name: userSnap.data()?.name || credential.user.displayName || "근로생",
      studentId
    };
  }

  state = await loadState();
  localStorage.setItem(SESSION_KEY, JSON.stringify({ studentId }));
}

function handleLocalAuth({ name, studentId, password }) {
  const users = readJson(USER_KEY, []);

  if (authMode === "signup") {
    if (users.some((user) => user.studentId === studentId)) {
      throw new Error("이미 가입된 학번입니다.");
    }
    currentUser = {
      id: `local-${studentId}`,
      name,
      studentId,
      password
    };
    users.push(currentUser);
    localStorage.setItem(USER_KEY, JSON.stringify(users));
    localStorage.setItem(SESSION_KEY, JSON.stringify({ studentId }));
    return;
  }

  const user = users.find((item) => item.studentId === studentId && item.password === password);
  if (!user) {
    throw new Error("가입 정보가 없거나 비밀번호가 다릅니다.");
  }
  currentUser = user;
  localStorage.setItem(SESSION_KEY, JSON.stringify({ studentId }));
}

async function logout() {
  if (services.mode === "firebase" && firebaseApi) {
    try {
      await firebaseApi.signOut(firebaseApi.auth);
    } catch (error) {
      console.warn("Firebase 로그아웃 실패", error);
    }
  }
  localStorage.removeItem(SESSION_KEY);
  currentUser = null;
  selectedCourseId = null;
  render();
}

function addShift(form) {
  const data = new FormData(form);
  const day = String(data.get("day") || "");
  const start = String(data.get("start") || "");
  const end = String(data.get("end") || "");
  const repeat = data.get("repeat") === "on";
  const member = getCurrentMember();

  if (!member || !day || !start || !end) return;
  if (minutesFromHHMM(end) <= minutesFromHHMM(start)) {
    render("종료 시간은 시작 시간보다 뒤여야 합니다.");
    return;
  }

  state.shifts.push({
    id: makeId("shift"),
    workerName: member.name,
    studentId: member.studentId,
    day,
    start,
    end,
    repeat
  });
  scheduleSave();
  render("근로 일정이 추가되었습니다.");
}

function deleteShift(id) {
  const shift = state.shifts.find((item) => item.id === id);
  if (!shift || !canManageShift(shift)) return;
  state.shifts = state.shifts.filter((item) => item.id !== id);
  scheduleSave();
  render();
}

function setupCourse(form) {
  const data = new FormData(form);
  const courseId = String(data.get("courseId") || "");
  const pointer = String(data.get("pointer") || "");
  if (!courseId || !pointer) return;

  state.courseStates[courseId] = {
    pointer,
    setupBy: currentUser.name,
    setupAt: nowIso(),
    recoveredBy: null,
    recoveredAt: null
  };
  scheduleSave();
  selectedCourseId = courseId;
  render("세팅 상태가 저장되었습니다.");
}

function recoverCourse(courseId) {
  const record = state.courseStates[courseId];
  if (!record) return;
  state.courseStates[courseId] = {
    ...record,
    recoveredBy: currentUser.name,
    recoveredAt: nowIso()
  };
  scheduleSave();
  selectedCourseId = courseId;
  render("회수 상태가 저장되었습니다.");
}

function resetCourse(courseId) {
  delete state.courseStates[courseId];
  scheduleSave();
  selectedCourseId = courseId;
  render("수업 기록이 초기화되었습니다.");
}

function addTodo(form) {
  const data = new FormData(form);
  const title = String(data.get("title") || "").trim();
  const body = String(data.get("body") || "").trim();
  const dueDate = String(data.get("dueDate") || "");
  if (!title || !body || !dueDate) return;

  state.todos.push({
    id: makeId("todo"),
    title,
    body,
    dueDate,
    done: false,
    createdAt: nowIso()
  });
  scheduleSave();
  render("업무가 추가되었습니다.");
}

function toggleTodo(id) {
  state.todos = state.todos.map((todo) => (todo.id === id ? { ...todo, done: !todo.done } : todo));
  scheduleSave();
  render();
}

function deleteTodo(id) {
  state.todos = state.todos.filter((todo) => todo.id !== id);
  scheduleSave();
  render();
}

function addMemo(form) {
  const data = new FormData(form);
  const text = String(data.get("text") || "").trim();
  const notice = data.get("notice") === "on";
  if (!text) return;

  state.memos.push({
    id: makeId("memo"),
    text,
    notice,
    author: currentUser.name,
    createdAt: nowIso()
  });
  scheduleSave();
  render("메모가 등록되었습니다.");
}

function deleteMemo(id) {
  state.memos = state.memos.filter((memo) => memo.id !== id);
  scheduleSave();
  render();
}

function addEvent(form) {
  const data = new FormData(form);
  const title = String(data.get("title") || "").trim();
  const date = String(data.get("date") || "");
  const body = String(data.get("body") || "").trim();
  if (!title || !date) return;

  state.events.push({
    id: makeId("event"),
    title,
    date,
    body,
    createdAt: nowIso()
  });
  calendarCursor = new Date(`${date}T00:00:00`);
  scheduleSave();
  render("일정이 추가되었습니다.");
}

function deleteEvent(id) {
  state.events = state.events.filter((event) => event.id !== id);
  scheduleSave();
  render();
}

async function reconcileCurrentMember(options = {}) {
  if (!currentUser) return null;

  const existing = state.members.find((member) => member.studentId === currentUser.studentId);
  if (existing) {
    currentUser = {
      ...currentUser,
      name: existing.name,
      role: existing.role,
      status: existing.status,
      color: existing.color
    };
    return existing;
  }

  const shouldApprove = shouldAutoApprove(currentUser.studentId);
  const member = {
    id: currentUser.id || `member-${currentUser.studentId}`,
    name: currentUser.name || "근로생",
    studentId: currentUser.studentId,
    role: shouldApprove ? "admin" : "worker",
    status: shouldApprove ? "approved" : "pending",
    color: WORKER_COLORS[state.members.length % WORKER_COLORS.length],
    createdAt: nowIso(),
    approvedAt: shouldApprove ? nowIso() : null,
    approvedBy: shouldApprove ? "system" : null,
    bannedAt: null
  };

  state.members.push(member);
  currentUser = { ...currentUser, role: member.role, status: member.status, color: member.color };
  await syncMemberProfile(member);
  scheduleSave();
  return member;
}

function shouldAutoApprove(studentId) {
  if (ADMIN_STUDENT_IDS.includes(studentId)) return true;
  return !state.members.some((member) => member.role === "admin" && member.status === "approved");
}

async function syncMemberProfile(member) {
  if (services.mode !== "firebase" || !firebaseApi) return;
  try {
    const userRef = firebaseApi.doc(firebaseApi.db, "users", member.studentId);
    await firebaseApi.setDoc(
      userRef,
      {
        name: member.name,
        studentId: member.studentId,
        role: member.role,
        status: member.status,
        color: member.color,
        updatedAt: firebaseApi.serverTimestamp()
      },
      { merge: true }
    );
  } catch (error) {
    console.warn("멤버 프로필 저장 실패", error);
  }
}

function getApprovedMembers() {
  return state.members.filter((member) => member.status === "approved");
}

function getMembersSorted() {
  const rank = { pending: 0, approved: 1, banned: 2 };
  return [...state.members].sort((a, b) => {
    return (rank[a.status] ?? 9) - (rank[b.status] ?? 9) || a.name.localeCompare(b.name, "ko");
  });
}

function getCurrentMember() {
  if (!currentUser) return null;
  return state.members.find((member) => member.studentId === currentUser.studentId) || null;
}

function isCurrentAdmin() {
  const member = getCurrentMember();
  return member?.status === "approved" && member?.role === "admin";
}

function updateMember(studentId, updater) {
  state.members = state.members.map((member) => {
    if (member.studentId !== studentId) return member;
    return updater(member);
  });
  const updated = state.members.find((member) => member.studentId === studentId);
  if (updated) {
    syncMemberProfile(updated);
  }
  if (currentUser?.studentId === studentId) {
    currentUser = {
      ...currentUser,
      role: updated.role,
      status: updated.status,
      color: updated.color
    };
  }
  scheduleSave();
}

function approveMember(studentId) {
  if (!isCurrentAdmin()) return;
  updateMember(studentId, (member) => ({
    ...member,
    status: "approved",
    role: member.role || "worker",
    approvedAt: nowIso(),
    approvedBy: currentUser.studentId,
    bannedAt: null
  }));
  render("근로생 가입을 승인했습니다.");
}

function banMember(studentId) {
  if (!isCurrentAdmin() || studentId === currentUser.studentId) return;
  updateMember(studentId, (member) => ({
    ...member,
    status: "banned",
    bannedAt: nowIso(),
    bannedBy: currentUser.studentId
  }));
  state.shifts = state.shifts.filter((shift) => shift.studentId !== studentId);
  scheduleSave();
  render("근로생을 명단에서 제외했습니다.");
}

function restoreMember(studentId) {
  if (!isCurrentAdmin()) return;
  updateMember(studentId, (member) => ({
    ...member,
    status: "approved",
    approvedAt: nowIso(),
    approvedBy: currentUser.studentId,
    bannedAt: null,
    bannedBy: null
  }));
  render("근로생 계정을 복구했습니다.");
}

function getPointerStatuses() {
  return [1, 2, 3, 4].map((no) => {
    const course = state.courses.find((item) => {
      const record = state.courseStates[item.id];
      return record?.pointer === String(no) && record?.setupAt && !record?.recoveredAt;
    });
    return { no, course };
  });
}

function getCourseStatus(courseId) {
  const record = state.courseStates[courseId];
  if (!record?.setupAt) {
    return { kind: "pending", label: "세팅 전", badgeClass: "rose" };
  }
  if (record.recoveredAt) {
    return { kind: "done", label: "회수 완료", badgeClass: "green" };
  }
  if (record.pointer === "none") {
    return { kind: "no-pointer", label: "포인터 없음", badgeClass: "teal" };
  }
  return { kind: "setup", label: `포인터 ${record.pointer}번 세팅`, badgeClass: "amber" };
}

function isCourseRecovered(courseId) {
  return Boolean(state.courseStates[courseId]?.recoveredAt);
}

function getShiftsForDay(day) {
  return state.shifts.filter((shift) => shift.day === day);
}

function getKoreanDay(date) {
  const jsDay = date.getDay();
  return DAYS.find((day) => DAY_INDEX[day] === jsDay) || "주말";
}

function canManageShift(shift) {
  return Boolean(currentUser?.studentId && shift.studentId === currentUser.studentId);
}

function findWorker(name, studentId = "") {
  return getApprovedMembers().find((worker) => worker.studentId === studentId) || getApprovedMembers().find((worker) => worker.name === name);
}

function sortByShiftTime(a, b) {
  return a.start.localeCompare(b.start) || a.workerName.localeCompare(b.workerName, "ko");
}

function sortByCourseTime(a, b) {
  return extractHour(a.time) - extractHour(b.time) || a.title.localeCompare(b.title, "ko");
}

function sortByDueDate(a, b) {
  return a.done - b.done || a.dueDate.localeCompare(b.dueDate) || a.createdAt.localeCompare(b.createdAt);
}

function sortByDateAsc(a, b) {
  return a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt);
}

function sortByCreatedDesc(a, b) {
  return b.createdAt.localeCompare(a.createdAt);
}

function renderTimeOptions(selected = "") {
  return Array.from({ length: 49 }, (_, index) => {
    const minutes = index * 30;
    const value = minutesToHHMM(minutes);
    return `<option value="${value}" ${value === selected ? "selected" : ""}>${value}</option>`;
  }).join("");
}

function minutesFromHHMM(value) {
  const [hour, minute] = String(value).split(":").map(Number);
  return (Number.isFinite(hour) ? hour : 0) * 60 + (Number.isFinite(minute) ? minute : 0);
}

function minutesToHHMM(minutes) {
  const hour = Math.floor(minutes / 60);
  const minute = minutes % 60;
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}

function extractHour(time) {
  const match = time.match(/(\d{1,2}):/);
  return match ? Number(match[1]) : 99;
}

function readJson(key, fallback) {
  try {
    const value = localStorage.getItem(key);
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function makeId(prefix) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function nowIso() {
  return new Date().toISOString();
}

function formatDateInput(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function formatKoreanDate(date) {
  return date.toLocaleDateString("ko-KR", {
    year: "numeric",
    month: "long",
    day: "numeric"
  });
}

function formatKoreanDateTime(value) {
  return new Date(value).toLocaleString("ko-KR", {
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function normalizeStudentId(value) {
  return value.replace(/\s+/g, "").toLowerCase();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll("`", "&#096;");
}
