# 나는 사회복지학전공 근로생이다

국립부경대학교 사회복지학전공 전공사무실 근로생 업무 관리 웹앱입니다.

## 포함 기능

- 학번 기반 로그인/회원가입
- 주간 근로생 시간표
- 2026-1학기 수업 세팅/회수 현황
- 포인터 1~4번 사용 상태
- 캘린더 일정
- To Do List
- 공지 고정 메모장
- Firebase 미설정 시 로컬 저장소 데모 모드

## 로컬 실행

이 폴더에서 Node.js가 설치되어 있다면 다음 명령으로 실행합니다.

```bash
node server.mjs
```

브라우저에서 `http://localhost:4173`을 엽니다.

## Firebase 연결

1. Firebase Console에서 새 프로젝트를 만듭니다.
2. Web App을 추가합니다.
3. Authentication에서 Email/Password 로그인을 활성화합니다.
4. Firestore Database를 생성합니다.
5. `firebase-config.example.js`의 값을 복사해서 `firebase-config.js`에 넣습니다.
6. `firestore.rules` 내용을 Firebase Firestore Rules에 배포합니다.

앱에서는 학번을 Firebase Auth 이메일로 변환해 사용합니다.

예시:

```text
202600000 -> 202600000@pknu-work.app
```

Firebase 웹 설정값은 비밀 키가 아닙니다. 실제 보안은 Firestore Rules와 Authentication으로 관리합니다.

## Vercel 배포

1. GitHub에 이 폴더를 저장소로 올립니다.
2. Vercel에서 `New Project`를 선택합니다.
3. GitHub 저장소를 연결합니다.
4. Framework Preset은 `Other`로 둡니다.
5. Build Command는 비워둡니다.
6. Output Directory도 비워둡니다.
7. 배포 후 발급된 URL에서 앱을 확인합니다.

## GitHub 작업 흐름

추천 흐름:

```bash
git init
git add .
git commit -m "Initial office work app"
git branch -M main
git remote add origin https://github.com/YOUR_NAME/YOUR_REPO.git
git push -u origin main
```

현재 PC에서 `git` 명령이 인식되지 않으면 Git for Windows 또는 GitHub Desktop을 설치한 뒤 진행합니다.

## 파일 구조

```text
index.html                 앱 진입점
styles.css                 화면 스타일
app.js                     앱 로직
firebase-config.js         Firebase 설정
firebase-config.example.js Firebase 설정 예시
firestore.rules            Firestore 보안 규칙
vercel.json                Vercel 설정
server.mjs                 로컬 정적 서버
```
