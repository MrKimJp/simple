// ──────────────────────────────────────────
// 시트 버튼용 폼 (기존 유지)
// ──────────────────────────────────────────
function showForm() {
  const html = HtmlService
    .createHtmlOutputFromFile("inputForm")
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "세부정보 입력으로 파일 생성");
}

// ──────────────────────────────────────────
// 웹앱 진입점
// ──────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("careLog")
    .setTitle("간병일지 입력")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ──────────────────────────────────────────
// 간병인 로그인 — 이름 + 생년월일로 인증
// 회원가입일자는 응답 시트 타임스탬프 자동 추출
// ──────────────────────────────────────────
function 간병인로그인(name, birthDate) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("간병인 등록 신청서(응답)");
  const data  = sheet.getDataRange().getValues();

  const parts      = birthDate.split("-");
  const inputBirth = parts[0].substring(2) + parts[1] + parts[2]; // "900115"

  for (let i = 1; i < data.length; i++) {
    const rowName  = String(data[i][3]).trim();
    const rowJumin = String(data[i][4]).trim().replace(/[^0-9]/g, "");
    const rowBirth = rowJumin.substring(0, 6);

    if (rowName === name.trim() && rowBirth === inputBirth) {
      const registerDate = Utilities.formatDate(
        new Date(data[i][0]),
        Session.getScriptTimeZone(),
        "yyyy.MM.dd"
      );
      return {
        memberNo:     String(data[i][1]),
        name:         rowName,
        registerDate: registerDate
      };
    }
  }
  throw new Error("이름 또는 생년월일이 일치하지 않습니다.");
}

// ──────────────────────────────────────────
// 간병일지 저장 (임시저장 / 최종제출)
// ──────────────────────────────────────────
function 간병일지저장(data, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("간병일지");
  if (!sheet) {
    sheet = ss.insertSheet("간병일지");
    sheet.appendRow([
      "제출일시", "회원번호", "간병인이름", "환자이름", "환자생년월일",
      "간병날짜", "시작시간", "종료시간",
      "식사보조", "활동보조", "배변보조", "위생보조", "기타",
      "특이사항", "결제금액", "상태"
    ]);
  }

  const now    = new Date();
  const status = (mode === "submit") ? "최종제출" : "임시저장";

  // 같은 회원번호 + 환자 조합의 기존 임시저장 행 삭제
  const allData = sheet.getDataRange().getValues();
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][1]) == String(data.memberNo) &&
        allData[i][3] == data.patientName &&
        allData[i][15] == "임시저장") {
      sheet.deleteRow(i + 1);
    }
  }

  data.logs.forEach(log => {
    sheet.appendRow([
      now,
      data.memberNo,
      data.caregiverName,
      data.patientName,
      data.patientBirth,
      log.date,
      log.startTime,
      log.endTime,
      log.식사보조 ? "O" : "",
      log.활동보조 ? "O" : "",
      log.배변보조 ? "O" : "",
      log.위생보조 ? "O" : "",
      log.기타     ? "O" : "",
      log.note,
      data.amount || "",
      status
    ]);
  });

  if (mode !== "submit") return "saved";
  return 전체문서생성(data);
}

// ──────────────────────────────────────────
// 최종제출: 세 문서 한 번에 생성
// ──────────────────────────────────────────
function 전체문서생성(data) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();

  // 드라이브 폴더 (기존 있으면 재사용)
  const 루트폴더   = DriveApp.getFolderById("1M6NjJ6bznAhdv1ohTFepvKRpNKB_47ty");
  const folderName = data.memberNo + "_" + data.caregiverName;
  let folder;
  const folders = 루트폴더.getFoldersByName(folderName);
  folder = folders.hasNext() ? folders.next() : 루트폴더.createFolder(folderName);

  // 간병 기간 문자열 조합 (날짜 오름차순 정렬)
  const sortedLogs = [...data.logs].sort((a, b) => a.date.localeCompare(b.date));
  const firstDate  = sortedLogs[0].date.replace(/-/g, ".");
  const lastDate   = sortedLogs[sortedLogs.length - 1].date.replace(/-/g, ".");
  const totalDays  = sortedLogs.length;
  const periodStr  = firstDate + " ~ " + lastDate + " (총 " + totalDays + "일)";

  // 1. 회원가입확인서
  _회원가입확인서채우기(data);
  SpreadsheetApp.flush();
  const 가입GID = ss.getSheetByName("2.회원가입확인서").getSheetId();
  PDF변환후저장(ssId, 가입GID,
    data.memberNo + "_" + data.caregiverName + "_회원가입확인서", folder);

  // 2. 간병인파견확인서
  _파견확인서채우기(data, periodStr);
  SpreadsheetApp.flush();
  const 파견GID = ss.getSheetByName("3.간병인파견확인서").getSheetId();
  PDF변환후저장(ssId, 파견GID,
    data.memberNo + "_" + data.caregiverName + "_간병인파견확인서", folder);

  // 3. 간병일지 (서명 이미지 삽입 때문에 flush + sleep)
  _간병일지출력채우기(data, firstDate, lastDate);
  SpreadsheetApp.flush();
  Utilities.sleep(2000);
  const 일지GID  = ss.getSheetByName("4.간병일지출력").getSheetId();
  PDF변환후저장(ssId, 일지GID,
    "간병일지_" + data.caregiverName + "_" + firstDate + "~" + lastDate, folder);

  return folder.getUrl();
}

// ──────────────────────────────────────────
// 회원가입확인서 시트 채우기
// ──────────────────────────────────────────
function _회원가입확인서채우기(data) {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName("간병인 등록 신청서(응답)");
  const templateSheet = ss.getSheetByName("2.회원가입확인서");
  const rows          = responseSheet.getDataRange().getValues();

  let foundRow = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) == String(data.memberNo)) {
      foundRow = rows[i];
      break;
    }
  }
  if (!foundRow) throw new Error("회원번호를 찾을 수 없습니다: " + data.memberNo);

  const 생년월일 = 주민번호변환(String(foundRow[4]));
  const 연락처   = foundRow[5];

  templateSheet.getRange("D7").setValue(data.registerDate);  // 응답 타임스탬프
  templateSheet.getRange("D8").setValue(data.caregiverName);
  templateSheet.getRange("D9").setValue(data.memberNo);
  templateSheet.getRange("D10").setValue(생년월일);
  templateSheet.getRange("D11").setValue(연락처);
}

// ──────────────────────────────────────────
// 간병인파견확인서 시트 채우기
// ──────────────────────────────────────────
function _파견확인서채우기(data, periodStr) {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName("간병인 등록 신청서(응답)");
  const templateSheet = ss.getSheetByName("3.간병인파견확인서");
  const rows          = responseSheet.getDataRange().getValues();

  let foundRow = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) == String(data.memberNo)) {
      foundRow = rows[i];
      break;
    }
  }
  if (!foundRow) throw new Error("회원번호를 찾을 수 없습니다: " + data.memberNo);

  const 간병인생년월일 = 주민번호변환(String(foundRow[4]));
  const 간병인연락처   = foundRow[5];
  const 병원          = foundRow[9];

  // 결제금액 처리
  const 숫자금액 = String(data.amount || "0").replace(/[^0-9]/g, "");
  const 원화금액 = Number(숫자금액).toLocaleString("ko-KR");
  const 한글금액 = 숫자한글변환(Number(숫자금액));
  const 표시금액 = 원화금액 + "원(" + 한글금액 + ")";

  templateSheet.getRange("D7").setValue(data.patientName);
  templateSheet.getRange("F7").setValue(data.patientBirth);
  templateSheet.getRange("D8").setValue(data.caregiverName);
  templateSheet.getRange("F8").setValue(간병인생년월일);
  templateSheet.getRange("D9").setValue(data.memberNo);
  templateSheet.getRange("F9").setValue(간병인연락처);
  templateSheet.getRange("C10").setValue(병원);
  templateSheet.getRange("C11").setValue(periodStr);   // 간병일지 날짜 자동 연동
  templateSheet.getRange("C12").setValue(표시금액);   // 간병일지 결제금액 자동 연동
}

// ──────────────────────────────────────────
// 간병일지 출력 시트 채우기
// ──────────────────────────────────────────
function _간병일지출력채우기(data, firstDate, lastDate) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("4.간병일지출력");
  if (!sheet) throw new Error("'4.간병일지출력' 시트가 없습니다.");

  const 간병인성별 = 성별변환(foundRow[4]);  // 간병인 주민번호
  const 환자성별   = 성별변환(foundRow[7]); // 환자 주민번호
  const 환자연락처 = foundRow[8];
  const 간병인연락처 = found [5];
  const 환자생년월일 = 주민번호변환(String(foundRow[4]));
  const 생년월일 = 주민번호변환(String(foundRow[4]));

  // 헤더
  sheet.getRange("B2").setValue("간  병  일  지");

  // 환자 정보
  templateSheet.getRange("C7").setValue(data.patientName);
  templateSheet.getRange("F7").setValue(환자연락처);
  templateSheet.getRange("C8").setValue(환자생년월일);
  templateSheet.getRange("F8").setValue(환자성별);

  // 간병인 정보
  templateSheet.getRange("C11").setValue(data.caregiverName);
  templateSheet.getRange("F11").setValue(간병인연락처);
  templateSheet.getRange("C12").setValue(간병인생년월일);
  templateSheet.getRange("F12").setValue(간병인성별);

  templateSheet.getRange("B16:H31").setValues([[
    "간병일자", "간병시간", "식사보조", "활동보조", "배변보조", "위생보조", "기타"
  ]]);

  // 날짜별 데이터
  data.logs.forEach((log, i) => {
    sheet.getRange(11 + i, 2, 1, 7).setValues([[
      log.date,
      log.startTime + " ~ " + log.endTime,
      log.식사보조 ? "O" : "",
      log.활동보조 ? "O" : "",
      log.배변보조 ? "O" : "",
      log.위생보조 ? "O" : "",
      log.기타     ? "O" : ""
    ]]);
    if (log.note) sheet.getRange(11 + i, 9).setValue("※ " + log.note);
  });

  // 동의 문구
  const consentRow = 11 + data.logs.length + 2;
  sheet.getRange(consentRow, 2, 1, 7).merge();
  sheet.getRange(consentRow, 2).setValue(
    "환자( " + data.patientName + " )는 간병인( " + data.caregiverName +
    " )으로부터 위와 같은 내용을 확인하고 서비스를 제공받았음을 동의합니다."
  );

  // 서명
  const signRow    = consentRow + 2;
  sheet.getRange(signRow, 6).setValue("환자 또는 보호자 서명:");
  const base64Data = data.signature.replace(/^data:image\/png;base64,/, "");
  const imgBlob    = Utilities.newBlob(
    Utilities.base64Decode(base64Data), "image/png", "signature.png"
  );
  const img = sheet.insertImage(imgBlob, 7, signRow);
  img.setWidth(104);
  img.setHeight(32);
}

// ──────────────────────────────────────────
// PDF 변환 공통
// ──────────────────────────────────────────
function PDF변환후저장(ssId, sheetGid, fileName, folder) {
  const url = "https://docs.google.com/spreadsheets/d/" + ssId +
    "/export?format=pdf&gid=" + sheetGid +
    "&size=A4&portrait=true&fitw=true" +
    "&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false";

  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token }
  });
  folder.createFile(response.getBlob().setName(fileName + ".pdf"));
}

// ──────────────────────────────────────────
// 유틸 함수
// ──────────────────────────────────────────
function 주민번호변환(jumin) {
  if (!jumin) return "";
  const front      = jumin.replace(/[^0-9]/g, "");
  const genderCode = front.substring(6, 7);
  const yearPrefix = (genderCode == "3" || genderCode == "4") ? "20" : "19";
  return yearPrefix + front.substring(0, 2) + "." +
         front.substring(2, 4) + "." + front.substring(4, 6);
}

function 숫자한글변환(num) {
  if (!num || num === 0) return "영원정";
  const units      = ["", "만", "억", "조"];
  const smallUnits = ["", "십", "백", "천"];
  const numbers    = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"];
  let result = "";
  let unitIndex = 0;

  while (num > 0) {
    let chunk = num % 10000;
    let chunkResult = "";
    let smallUnitIndex = 0;
    while (chunk > 0) {
      const digit = chunk % 10;
      if (digit !== 0) {
        chunkResult = numbers[digit] + smallUnits[smallUnitIndex] + chunkResult;
      }
      chunk = Math.floor(chunk / 10);
      smallUnitIndex++;
    }
    if (chunkResult) result = chunkResult + units[unitIndex] + result;
    num = Math.floor(num / 10000);
    unitIndex++;
  }
  return result + "원정";
}

function 성별변환(jumin) {
  if (!jumin) return "";
  const genderCode = String(jumin).replace(/[^0-9]/g, "").substring(6, 7);
  if (genderCode == "1" || genderCode == "3") return "남";
  if (genderCode == "2" || genderCode == "4") return "여";
  return "";


// ──────────────────────────────────────────
// 기존 제출 기록 불러오기 (수정 모드용)
// ──────────────────────────────────────────
function 기존간병일지불러오기(memberNo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("간병일지");
  if (!sheet) return null;

  const allData = sheet.getDataRange().getValues();
  const logs    = [];

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (String(row[1]) == String(memberNo) && row[15] == "최종제출") {
      logs.push({
        date:      row[5],
        startTime: row[6],
        endTime:   row[7],
        식사보조:  row[8]  == "O",
        활동보조:  row[9]  == "O",
        배변보조:  row[10] == "O",
        위생보조:  row[11] == "O",
        기타:      row[12] == "O",
        note:      row[13],
        // 공통 정보 (첫 번째 행에서)
        patientName:  row[3],
        patientBirth: row[4],
        amount:       row[14]
      });
    }
  }

  if (logs.length === 0) return null;

  return {
    patientName:  logs[0].patientName,
    patientBirth: logs[0].patientBirth,
    amount:       logs[0].amount,
    logs: logs.map(l => ({
      date:      l.date,
      startTime: l.startTime,
      endTime:   l.endTime,
      식사보조:  l.식사보조,
      활동보조:  l.활동보조,
      배변보조:  l.배변보조,
      위생보조:  l.위생보조,
      기타:      l.기타,
      note:      l.note
    }))
  };
}

// ──────────────────────────────────────────
// 수정 제출 — 기존 데이터/PDF 삭제 후 재생성
// ──────────────────────────────────────────
function 간병일지수정(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 시트에서 기존 최종제출 행 삭제
  const sheet   = ss.getSheetByName("간병일지");
  const allData = sheet.getDataRange().getValues();
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][1]) == String(data.memberNo) &&
        allData[i][3] == data.patientName) {
      sheet.deleteRow(i + 1);
    }
  }

  // 2. 드라이브에서 기존 PDF 삭제
  const 루트폴더   = DriveApp.getFolderById("1M6NjJ6bznAhdv1ohTFepvKRpNKB_47ty");
  const folderName = data.memberNo + "_" + data.caregiverName;
  const folders    = 루트폴더.getFoldersByName(folderName);
  if (folders.hasNext()) {
    const folder = folders.next();
    const files  = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      // 간병일지, 파견확인서, 가입확인서 PDF 삭제
      if (name.endsWith(".pdf")) {
        file.setTrashed(true);
      }
    }
  }

  // 3. 새 데이터로 저장 + PDF 재생성
  return 간병일지저장(data, "submit");
}
