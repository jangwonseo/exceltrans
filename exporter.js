
/**
아래는 exceljs에서 issue로 나온 IE11 사용시 에러에 대한 해결방법 코드인데 IPCMS에서는 해결안됨
(exceljs 임포트 전에 반드시 먼저 선언할 것.)

const rewritePattern = require('regexpu-core');
const {generateRegexpuOptions} = require('@babel/helper-create-regexp-features-plugin/lib/util');

const {RegExp} = window;
try {
  new RegExp('a', 'u');
} catch (err) {
    window.RegExp = function(pattern, flags) {
    if (flags && flags.includes('u')) {
      return new RegExp(rewritePattern(pattern, flags, generateRegexpuOptions({flags, pattern})));
    }
    return new RegExp(pattern, flags);
  };
  window.RegExp.prototype = RegExp;
}
import ExcelJs from "exceljs/dist/exceljs";


****** 문제 및 해결 상태 : 
1. IPCMS에 적용된 exceljs는 4.1.1 버전이었음
2. 4.1.1 버전에서 IE11 일경우 구문에러(syntax Error) 발생
3. vue-project에서 2.0.1 버전일경우 IE11에서도 에러는 발생하지않지만,
   IPCMS 에서는 2.0.1 버전으로 Downgrade해도 IE11에서 에러 발생.
4. 하지만, 2.0.1 버전에서 exceljs 의 dist 파일을 임포트 하면 IE11에서 에러발생하지않음
5. 현재 2.0.1 버전으로 Downgrade 해놓고 dist파일을 임포트하여 임시로 동작하게 함
6. 추측 : IE11에서 스크립트소스 참조시, ES6->ES5 트랜스파일링하는 시점이 맞지않아, 구문에러발생

* vue-project에서 2.0.1 버전일경우, IE11에서 왜 에러가 발생하지않는지?
  (하지만 vue-project 에서도 4.1.1 버전으로 변경시, IE11에서 에러발생함.
   그리고 4.1.1 버전에서는 excel-js에서 참조하는 dependency 라이브러리중 트랜스파일링이 제대로 이루어지지않는 문제발생.
   그래서 vue-project 를 4.1.1 버전으로 upgrade 할수는 없는 상황.)

* IPCMS에서 4.1.1 버전을 사용하면서 exceljs 의 dist 파일을 임포트해도 에러발생함
  즉, 4.1.1 버전에서는 dependency 라이브러리가 변경되었고, 변경된 하위의 dependency 라이브러리에서 es6가 트랜스파일링이
  안된다고 추측중

*/

 /**
 * json 데이터로 엑셀파일 생성
 * 
 *  엑셀 스타일을 지정하고, 엑셀 데이터를 채우는 동작을 수행
 *  사용법은 listGpoaByAttorney.js 참고
 * 
 * created by jangwon.seo
 */
// import ExcelJs from "exceljs"; // 원래 정상 작동코드이지만 현재 임시로 dist 파일을 import 하도록 변경
import ExcelJs from "exceljs/dist/exceljs";

export const ExcelExporter = {

    workbook: null,

    // 초기화
    init: function(firstSheetName) {
    	const userId = sessionStorage.getItem("userId"); // 세션스토리지에 저장된 userId
        this.workbook = new ExcelJs.Workbook();
        this.workbook.creator = userId;
        this.workbook.lastModifiedBy = userId;
        this.workbook.addWorksheet(firstSheetName); // worksheet 객체 생성
    },
    
    // 엑셀에 출력할 데이터 세팅 함수
    setData: function(jsonData) {
        
        const worksheet = this.workbook.getWorksheet(jsonData.sheetName);
        const rowList = jsonData.rowList; 
        
        // 테이블 헤더(thead) 셋팅(First Row)
        const thead = rowList[0].map(cell => {
            return { key: cell.address, header: cell.value };
        });
        worksheet.columns = thead;
    
        // 테이블 내용(tbody) 셋팅(Except For First Row)
        let hasImage = false;
        const tbody = rowList.slice(1, rowList.length).map((row) => {
            return row.map((cell) => {
                if (typeof cell.value === "object") {   // 이미지 객체인경우
                    hasImage = true;
                    return "";
                } else {
                    return cell.value;
                }
            });    
        });
        worksheet.addRows(tbody);

        // 테이블 내용(tbody) 중, 이미지가 있는 경우 이미지 셋팅(반드시, 데이터세팅후 이미지세팅 가능)
        if (hasImage) { 
            rowList.slice(1, rowList.length).forEach((row, rowIndex) => {
                row.forEach((cell, colIndex) => {
                    if (typeof cell.value === "object") {   // 이미지 객체인경우
                        setImageToCell(worksheet, rowIndex, colIndex, cell.value); 
                    }
                });    
            });
        }
    },

    // 엑셀에 출력할 스타일 세팅 함수
    setStyle: function(jsonData) {

        const sheetName = jsonData.sheetName;
        let worksheet = this.workbook.getWorksheet(sheetName);
        let styleObj = jsonData.sheetStyle;
        let worksheetThead = worksheet.getRow(1);

        // style 객체 초기화
        if (this.isEmptyObject(styleObj)) styleObj = { theme: "default", thead: {} };

        // 테마별 지정해둔 스타일정보 
        if (styleObj.theme === "table") {
            styleObj = getTableThemeStyle();
        } else if (styleObj.theme === "custom") {
            styleObj = customThemeStyle(styleObj);
            if (styleObj.thead == null) styleObj = getDefaultThemeStyle();  // custom 추가
        } else {    // default
            styleObj = getDefaultThemeStyle();
        }

        ///////////////////////// worksheet 객체에 style 정보 세팅 //////////////////////////////////

        // table 과 custom 일때만 조작하는 스타일 정보
        if (["table", "custom"].includes(styleObj.theme)) {
            
            // thead 높이 지정
            worksheetThead.height = styleObj.thead.height;

            // thead cell 지정
            worksheetThead.eachCell({ includeEmpty: true }, cell => {
                cell.border = styleObj.thead.border;
                cell.font = styleObj.thead.font;
                cell.alignment = styleObj.thead.alignment;
                cell.fill = styleObj.thead.fill;
            });
        }

        // custom 일때만 조작하는 스타일 정보
        if (["custom"].includes(styleObj.theme)) {
            // thead 너비 지정
        	if (styleObj.thead.width) {
        		styleObj.thead.width.forEach(col => worksheet.getColumn(col.address).width = col.width);
        	} else {	// thead 너비 스타일이 없는 경우, auto width 처리
        		worksheet.columns.forEach(function (column, i) {
        			let maxLength = 0;
        			column["eachCell"]({ includeEmpty: true }, function (cell) {
        				let columnLength = cell.value ? cell.value.toString().length : 15;
        				if (columnLength > maxLength) {
        					maxLength = columnLength;
        				}
        			});
        			column.width = maxLength < 15 ? 15 : maxLength > 200 ? 100 : maxLength + 5;	// max size(200)을 초과하는 경우
        		});
        	}
            
            // 병합 지정
            if (styleObj.mergeCellList) {
                styleObj.mergeCellList.forEach(range => {
                    worksheet.mergeCells(range);
                });
            }
        }
    
        // tbody 윤곽선 지정
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            if (rowNumber != 1) {
                row.eachCell({ includeEmpty: true }, cell => {
                    cell.border = staticValue.border
                });
            } else {
                if ("default" === styleObj.theme) {
                    row.eachCell({ includeEmpty: true }, cell => {
                        cell.border = staticValue.border
                    });
                }
            }
        });

        // 전체 행 수직 중앙 정렬
        worksheet.eachRow({ includeEmpty: true }, (row) => {
            row.eachCell({ includeEmpty: true }, cell => {
                if (!cell.alignment) {
                    cell.alignment = { vertical: "middle", wrapText: true };
                }
            });
        });
    },

    // 엑셀에 출력할 이미지 수동 세팅 함수(원하는 위치에 직접 세팅)
    setImage: function(jsonData) {
        let worksheet = this.workbook.getWorksheet(jsonData.sheetName);

        jsonData.imageList.forEach((image) => {
            const imageId = this.workbook.addImage(image.data);
            worksheet.addImage(imageId, image.range);
        });
    },
    
    excelFileName: function(fileName) {
      return fileName == null
        ? "Excel.xlsx"
        : fileName.includes(".xls")
        ? fileName
        : fileName + ".xlsx";
    },

    isEmptyObject: function(obj) {
        if (obj == undefined) obj = {};
        return (
            Object.keys(obj).length === 0 &&
            JSON.stringify(obj) === JSON.stringify({})
        );
    }
};

// 기본 테마 
const getDefaultThemeStyle = () => { 
    return { theme: "default", thead: staticValue.theme.defaultThead }; 
};

// 테이블 테마
const getTableThemeStyle = () => { 
    return { theme: "table", thead: staticValue.theme.tableThead }; 
};

// 커스텀(Custom) 테마
const customThemeStyle = (styleObj) => {
    return styleObj;
};


// 고정값
const staticValue = {

    theme: {
        // default 테마용 스타일
        defaultThead: {
            heigth: 15,
            width: []
        },

        // table 테마용 스타일
        tableThead: {
            heigth: 30,
            width: [],
            fill: {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "AADBFF" }
            },
            alignment: {
                vertical: "middle",
                horizontal: "center"
            },
            font: {
                size: 13,
                bold: true,
                family: 2
            },
            border: {
                top: { style: "thin" },
                left: { style: "thin" },
                right: { style: "thin" },
                bottom: { style: "thin" }
            }
        }
    },

    // 기본 윤곽선 스타일
    border: {
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
        bottom: { style: "thin" }
    }
};

/**
 * Cell 이미지 세팅 함수
 * 
 * @param {*} worksheet 
 * @param {*} rowIndex  - 행 인덱스
 * @param {*} colIndex  - 열 인덱스
 * @param {*} imageObj - { base64:(...), extension:(...), width:(...), height:(...) }
 */
const setImageToCell = (worksheet, rowIndex, colIndex, imageObj) => {   
    
    const imageId = worksheet._workbook.addImage({ 
        base64: imageObj.base64, 
        extension: imageObj.extension 
    });
    
    worksheet.addImage(imageId, {
        tl: { col: colIndex + 0.1, row: (rowIndex+1) + 0.1 },	    // 이미지 위치 index(이미지를 중앙에 위치하기 위해 0.1씩 이동)
        ext: { width: imageObj.width, height: imageObj.height },	// 이미지 사이즈
        editAs: 'oneCell'
    });

    // 이미지 크기에 맞게 column, row 사이즈 조정(기존 이미지 사이즈에 버퍼 공간 2px 추가)
    worksheet.getColumn(colIndex).width = (imageObj.width / 8) + 10;
    worksheet.getRow(rowIndex+2).height = imageObj.height / 1.3 < 17 ? 17 : (imageObj.height / 1.3) + 10;
};

const startLoadingBar = () => {
    if ($("div.ajax-loading-spinner").length < 1) {
        const $loadingBar = $('<div class="ajax-loading-spinner"><div class="inner"><div></div></div><span>Loading..</span></div>');
        $(document.body).append($loadingBar);
    }
}

const stopLoadingBar = () => {
    if ($("div.ajax-loading-spinner").length > 1) {
        $("div.ajax-loading-spinner").remove();
    }
}