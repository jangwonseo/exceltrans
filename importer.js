import ExcelJs from "exceljs";

/**
 * 엑셀파일로부터 데이터 추출
 * 
 * 사용법
 * init() 함수를 호출하여 추출하고자하는 데이터의 기준 위치를 지정한 후,
 * extractData() 함수를 호출하여 데이터를 추출한다.
 * 
 * 
 * --tip. 업로드용 엑셀파일의 0번째 행에 영어 변수명을 넣고 높이를 0으로 조절 후,
 *        1번째 행에 열 항목(Title) 이름(한글)을 지정하여 headerRowIndex=0 으로 셋팅한채로 
 *        데이터를 추출하면 영어변수명으로 이루어진 json 형태의 데이터로 추출할 수 있다.
 * 
 * created by jangwon.seo
 */

export const ExcelImporter = {

  workbook: null,
  headerRowIndex: 0,
  startBodyRowIndex: 1,

  /**
   * 초기화 - 추출 시작점(기준점) 지정
   * 
   * headerRowIndex - 추출할 엑셀의 열 항목(Title)의 행 인덱스(default: 0)
   * startBodyRowIndex - 추출할 엑셀의 데이터 행 시작 인덱스(default: 1)
   */
  init: function (headerRowIndex = 0, startBodyRowIndex = 1) {

    this.workbook = new ExcelJs.Workbook();

    this.headerRowIndex = headerRowIndex || 0;
    this.startBodyRowIndex = startBodyRowIndex || 1;
    if (headerRowIndex >= startBodyRowIndex) {
      throw new Error('인덱스가 잘못되었습니다.');
    }
  },

  /**
   * 데이터 추출
   * 
   * file - file 객체
   */
  extractData: function (file) { // read async 
    return new Promise((resolve, reject) => {

      startLoadingBar();  // 로딩바 설정

      let reader = new FileReader();
      reader.readAsArrayBuffer(file);
      reader.onload = (e) => {           // onload() called after file read success

        const arrBuffer = e.target.result;

        this.workbook.xlsx
          .load(arrBuffer, { type: "array", cellDates: true, dateNF: "YYYY-MM-DD" })
          .then(wb => {
            let bodyRowList = [];   // 행
            let headerList = [];    // 열
            const firstSheet = wb.worksheets[0];    // 첫번째 Sheet 로 제한(첫Sheet만 읽음)
            firstSheet.eachRow({ includeEmpty: false }, (row, rowCount) => {
              let rowObj = {};
              row.eachCell({ includeEmpty: true }, (cell, cellCount) => {
                if (rowCount === this.headerRowIndex + 1) {
                  headerList.push(cell.value);
                } else if (rowCount > this.startBodyRowIndex) {

                  let cellValue = "";
                  if (Enums.ValueType.RichText === cell.model.type) {         // RichText 형인 경우
                    const value = cell.value.richText.map(richText => (richText.text || "").trim()).join("");
                    cellValue = value || null;
                  } else if (Enums.ValueType.Date === cell.model.type) {      // Date 형인경우
                    const value = (new Date(cell.value - cell.value.getTimezoneOffset() * 60000)).toISOString().split("T")[0];
                    cellValue = value || null;
                  } else {
                    cellValue = cell.value ? String(cell.value).trim() : null;
                  }

                  rowObj[headerList[cellCount - 1]] = cellValue;
                }
              });

              if (rowCount > this.startBodyRowIndex) {
                bodyRowList.push(rowObj);
              }
            });

            return { headerList, bodyRowList };
          })
          .then((data) => {
            const headerList = data.headerList;
            const bodyRowList = data.bodyRowList;

            resolve({ headerList, bodyRowList });   // 성공시 반환
          })
          .catch(err => {
            reject(new Error("Request is failed" + err));
          })
          .finally(() => {
            stopLoadingBar();
          });
      };

      reader.onerror = reject; // 읽기 에러발생시 실행
    });
  }
}

const startLoadingBar = () => {
  if ($("div.ajax-loading-spinner").length < 1) {
    const $loadingBar = $('<div class="ajax-loading-spinner"><div class="inner"><div></div></div><span>Loading..</span></div>');
    $(document.body).append($loadingBar);
  }
}

const stopLoadingBar = () => {
  if ($("div.ajax-loading-spinner").length > 0) {
    $("div.ajax-loading-spinner").remove();
  }
}

const Enums = {

  ValueType: {
    Null: 0,
    Merge: 1,
    Number: 2,
    String: 3,
    Date: 4,
    Hyperlink: 5,
    Formula: 6,
    SharedString: 7,
    RichText: 8,
    Boolean: 9,
    Error: 10,
  },
  FormulaType: {
    None: 0,
    Master: 1,
    Shared: 2,
  },
  RelationshipType: {
    None: 0,
    OfficeDocument: 1,
    Worksheet: 2,
    CalcChain: 3,
    SharedStrings: 4,
    Styles: 5,
    Theme: 6,
    Hyperlink: 7,
  },
  DocumentType: {
    Xlsx: 1,
  },
  ReadingOrder: {
    LeftToRight: 1,
    RightToLeft: 2,
  },
  ErrorValue: {
    NotApplicable: '#N/A',
    Ref: '#REF!',
    Name: '#NAME?',
    DivZero: '#DIV/0!',
    Null: '#NULL!',
    Value: '#VALUE!',
    Num: '#NUM!',
  }
}
