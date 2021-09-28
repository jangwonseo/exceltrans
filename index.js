import { ExcelImporter } from "./importer"
import { ExcelExporter } from "./exporter"

/**
 * 엑셀파일로부터 데이터 추출
 * 
 * 사용법
 * init() 함수를 호출하여 추출하고자하는 데이터의 기준 위치를 지정한 후,
 * extractData() 함수를 호출하여 데이터를 추출한다.
 * 
 * listSample.js 의 #excelImport 부분 코드를 참고할 수 있다.
 * 
 * --tip. 업로드용 엑셀파일의 0번째 행에 영어 변수명을 넣고 높이를 0으로 조절 후,
 *        1번째 행에 열 항목(Title) 이름(한글)을 지정하여 headerRowIndex=0 으로 셋팅한채로 
 *        데이터를 추출하면 영어변수명으로 이루어진 json 형태의 데이터로 추출할 수 있다.
 * 
 * created by jangwon.seo
 */

// import ExcelJs from "exceljs"; exceljs 버전문제로 dist 파일 참조. exporter.js 주석 참고

export const ExcelImporter;
export const ExcelExporter;

