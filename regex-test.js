const r1 = /(?:<EXCEL_CMD>|EXCEL_CMD>|CEL_CMD>)([\s\S]*?)<\/EXCEL_CMD>/gi;
const r2 = /(?:<EXCEL_CMD>|EXCEL_CMD>|CEL_CMD>)[\s\S]*$/gi;
const text = 'Hello world! <EXCEL_CMD>{"action": "add_sheet"}</EXCEL_CMD> and then CEL_CMD>{"action": "add_sheet"}</EXCEL_CMD> and finally an unclosed EXCEL_CMD>{"action';
console.log("CLEANED:", text.replace(r1, '').replace(r2, '').trim());
