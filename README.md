# Preposition Room · 전치사 3D 방

고등학교 영어 수업용 웹앱이에요. 학생이 영어 문장(예: `Put the lamp on the desk.`)을 쓰면 3D 방 안의 물건이 그 뜻대로 움직여요.

- **학생용 링크:** https://ereyang.github.io/preposition-room/
- **Free 자유 모드:** 침대·옷장·책상·책장이 놓인 방에 물건을 배치하고, 📸 Snap 찰칵으로 완성한 방을 선생님께 보내요.
- **Mission 미션 모드:** 목표 그림과 똑같이 방을 꾸미는 미션 6개.

## 파일
- `index.html` — 앱 전체 (HTML 파일 하나, Three.js r128 CDN)
- `gas/Code.gs` — 사진을 선생님 구글 드라이브 `전치사 방 앨범/학년 반` 폴더에 저장하는 Apps Script 웹 앱

## 사진 저장 (구글 Apps Script)
1. script.google.com 새 프로젝트에 `gas/Code.gs`를 붙여 넣고 `setup` 실행 → 권한 허용
2. 배포 → 새 배포 → 웹 앱 (실행: 나, 액세스: 모든 사용자)
3. 웹 앱 주소를 `index.html`의 `PREP_API`에 넣기
4. 코드를 고치면 배포 관리 → 수정 → 새 버전 (주소는 그대로)
