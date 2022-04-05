import * as xlsx from 'XLSX';

function readExcel() {
    const test = xlsx.readFile(__dirname + '/' + 'test2.xlsx', {type: "array"});
    const first_sheet_name = test.SheetNames[0];
    const worksheet = test.Sheets[first_sheet_name];
    // console.log(test)
    // console.log(test.SheetNames);
    // console.log(test.Sheets[first_sheet_name])
    // console.log(worksheet["!merges"]);
    // console.log(worksheet["!rows"]);
    // console.log(worksheet["!cols"]);
    // console.log(worksheet)
    // console.log(worksheet["!merges"])
    // Object.entries(worksheet).forEach(([cellName, value])=>{
    //     if(/^[A-Z]{1,3}[0-9]{1,7}$/g.test(cellName)){
    //         console.log("cell num : ",cellName,"value : ", value);
    //     }
    // })
    // console.log(worksheet)
    // console.log(worksheet["!rows"]);
    console.log(xlsx.utils.sheet_to_json(worksheet, {header:1}))
}

export function readCells(inputFileName: string) {
    const excel = xlsx.readFile(__dirname + '/' + inputFileName, {type: "array"});
    const first_sheet_name = excel.SheetNames[0];
    const worksheet = excel.Sheets[first_sheet_name];
    const data = xlsx.utils.sheet_to_json(worksheet,{header:1});
    console.log(data);
}

export function makeCells() {
    const workbook:xlsx.WorkBook = xlsx.utils.book_new();
    const worksheet:xlsx.WorkSheet = xlsx.utils.aoa_to_sheet([[1,2,3],[4,5,6],[7,8,9]])
    console.log(workbook);
    console.log(worksheet);
    console.log(xlsx.utils.sheet_to_json(worksheet, {header:1}))
}

export default readExcel;