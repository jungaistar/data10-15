// 로그인 설정 — 이 파일 한 곳만 바꾸면 된다(생성: data10-00/scripts/apply-auth-kit.mjs, 대문 문구 출처 docs/plan.json).
// 내 Supabase 로 옮길 때: url · anonKey 를 내 프로젝트 값으로, 카카오를 설정하지 않았으면 kakao: false. 안내 = docs/내_깃허브로_옮기기.md
// prefix 는 supabase/auth.sql 의 표 이름 앞부분과 같아야 한다.
window.AUTH_CONFIG = {
  "url": "https://hcmgdztsgjvzcyxyayaj.supabase.co",
  "anonKey": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhjbWdkenRzZ2p2emN5eHlheWFqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE0MzU4ODcsImV4cCI6MjA4NzAxMTg4N30.gznaPzY1l8qDAPsEyYNR9KS7f7VqS3xaw-_2HTSwSZw",
  "prefix": "data1015",
  "kakao": true,
  "idleMinutes": 30,
  "localBypass": true,
  "site": {
    "title": "합의서 · 인보이스 자동 작성",
    "summary": "협력사와 보상 합의를 마친 결과 엑셀을 올리면 협력사별로 묶어 금액을 검증하고, 회사 합의서 · 인보이스 엑셀 양식의 {{키}} 칸을 채워 서식 그대로의 엑셀과 인쇄용 PDF 를 만드는 도구입니다.",
    "course": "현장 데이터 디지털화 전문가과정 · 2차수",
    "features": [
      {
        "name": "합의 결과 가져오기",
        "desc": "엑셀 여러 개를 읽고 열 이름이 달라도 찾아 연결합니다. 협력사별로 건수 · 합의합계 · 부가세 · 총액을 묶습니다."
      },
      {
        "name": "금액 검증",
        "desc": "합의금액이 청구보다 크거나 비율 계산과 다른 줄을 「확인 필요 · 주의」로 알리고 그 협력사는 문서 만들기에서 뺍니다."
      },
      {
        "name": "양식 그대로 채우기",
        "desc": "회사 엑셀 양식의 {{키}} 칸만 채워 서식 · 병합 · 직인 그림을 그대로 둡니다. 표 줄이 모자라면 끼웁니다."
      },
      {
        "name": "엑셀 · PDF 내려받기",
        "desc": "합의서 · 인보이스를 엑셀(여러 개는 zip 한 파일)과 인쇄용 PDF 로 받습니다."
      },
      {
        "name": "문서 번호 · 발행 이력",
        "desc": "종류 · 달마다 일련번호, 같은 내용은 재발행(같은 번호), 바뀌면 새 번호와 「대체됨」."
      },
      {
        "name": "설정 · 백업",
        "desc": "발행처 · 입금 계좌 · 통화 · 부가세 · 반올림 · 허용 차이 · 번호 형식 · 합의 조건 문구, JSON 백업."
      }
    ],
    "steps": [
      "합의 결과 넣기",
      "협력사별 묶기",
      "금액 검증",
      "양식 고르기",
      "문서 만들기",
      "엑셀 · PDF 받기"
    ],
    "flow": [
      {
        "step": "합의 결과 넣기",
        "desc": "엑셀 — 머리행 · 열 이름 찾기, 못 찾은 열은 화면에서 연결"
      },
      {
        "step": "협력사별 묶기",
        "desc": "협력사 코드 우선으로 건수 · 청구 · 합의 · 부가세 · 총액"
      },
      {
        "step": "금액 검증",
        "desc": "합의금액은 칸 우선, 비면 청구 × 부담 비율 — 확인 필요 · 주의"
      },
      {
        "step": "양식 고르기",
        "desc": "기본 양식 또는 {{키}} 를 적은 회사 양식"
      },
      {
        "step": "문서 만들기",
        "desc": "키 칸 채우기 · 반복 줄 · 합계 범위, 문서 번호 확정"
      },
      {
        "step": "엑셀 · PDF 받기",
        "desc": "엑셀(zip) · 인쇄용 PDF, 발행 이력"
      }
    ],
    "goals": [
      "합의 결과 엑셀을 올리면 협력사 · 클레임별 합의금액을 자동으로 읽습니다.",
      "합의금액 · 비율 · 번호를 검증해 잘못된 줄을 문서를 만들기 전에 알립니다.",
      "회사가 쓰던 엑셀 양식의 서식을 그대로 두고 값만 채워 합의서 · 인보이스를 만듭니다."
    ],
    "screen": "docs/img/screen.png",
    "plan": "plan.html",
    "home": true,
    "landing": {
      "colors": {
        "navy": "#1a2440",
        "navy2": "#11192e",
        "brand": "#2c58c4",
        "brandDark": "#2347a0",
        "accent": "#9db8ff"
      }
    }
  }
};
