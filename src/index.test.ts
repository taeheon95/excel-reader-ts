import readExcel, {makeCells} from "./index";
import {readCells} from "./index";
import * as fs from 'fs'
import * as XLSX from 'xlsx';

describe('읽기', function () {
    it("test read json", () => {
        const fileJson = fs.readFileSync(__dirname + '/' + "test.json", {encoding: 'utf8'});
        const applyObject = JSON.parse(fileJson);
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet(applyObject);
        console.log()
        XLSX.utils.book_append_sheet(workbook, worksheet, "test");
        // console.log(applyObject);
        console.log(worksheet);
        XLSX.writeFile(workbook, "applyTest.xlsx");
    })
    it("전체 객체", () => {
        readExcel();
    })

    it("셀 읽기", () => {
        readCells("test3.xlsx");
    })

    it("엑셀 to json array", () => {
        const workbook = XLSX.readFile(__dirname+"/../"+"applyTest.xlsx", {type:'array'});
        const firstSheet = workbook.SheetNames[0];
        const workSheet = workbook.Sheets[firstSheet];
        const sheetJson = XLSX.utils.sheet_to_json(workSheet);
        console.log(sheetJson);
    })

    it("엑셀 테스트", () => {
        makeCells()
    })
});