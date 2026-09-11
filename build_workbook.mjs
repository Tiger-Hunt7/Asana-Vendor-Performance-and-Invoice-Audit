import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const root = "/workspace/scratch/1ed9c5d2dd7d";
const repo = `${root}/Asana-Vendor-Performance-and-Invoice-Audit`;
const data = JSON.parse(await fs.readFile(`${root}/.asana_audit_work/sanitized_source.json`, "utf8"));
await fs.mkdir(repo, { recursive: true });

const wb = Workbook.create();
const dashboard = wb.worksheets.add("Dashboard");
const yearly = wb.worksheets.add("Yearly Summary");
const monthly = wb.worksheets.add("Monthly Summary");
const financial = wb.worksheets.add("Financial Summary");
const ledger = wb.worksheets.add("Financial Ledger");
const quality = wb.worksheets.add("Data Quality");
const tasks = wb.worksheets.add("Asana Tasks");

const navy = "#17365D", blue = "#2F75B5", pale = "#EAF3F8", gold = "#D6A84B";
const yellow = "#FFF2CC", white = "#FFFFFF", text = "#1F2937", red = "#F4CCCC", green = "#D9EAD3";
const header = { fill: navy, font: { bold: true, color: white }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
const section = { fill: blue, font: { bold: true, color: white }, verticalAlignment: "center" };
const money = '"$"#,##0';

function title(ws, range, value, subtitleRange, subtitle) {
  ws.showGridLines = false;
  ws.getRange(range).merge();
  ws.getRange(range.split(":")[0]).values = [[value]];
  ws.getRange(range).format = { font: { bold: true, size: 15, color: navy } };
  ws.getRange(subtitleRange).merge();
  ws.getRange(subtitleRange.split(":")[0]).values = [[subtitle]];
  ws.getRange(subtitleRange).format = { font: { italic: true, color: "#666666" }, wrapText: true };
}

// ASANA TASK DATA
title(tasks, "A1:N1", "FICTIONAL ASANA TASK EXPORT", "A2:N2", "147-record portfolio demonstration. Identities, locations, task IDs, links, and narrative details were replaced; analytical dates, status counts, and numeric fields were preserved.");
tasks.getRange("A4:N4").values = [["Task ID", "Task Name", "Reporting Date", "Year", "Month", "Project Section", "Status", "Completed", "Assignee", "Cost Field", "Created At", "Completed At", "Last Modified", "Notes"]];
tasks.getRange("A5:N151").values = data.tasks.map(r => [r[0], r[1], r[2] ? new Date(r[2]) : null, null, null, r[3], r[4], r[5], r[6], r[7], r[8] ? new Date(r[8]) : null, r[9] ? new Date(r[9]) : null, r[10] ? new Date(r[10]) : null, r[11]]);
tasks.getRange("D5:D151").formulasR1C1 = [["=IF(RC[-1]=\"\",\"Undated\",YEAR(RC[-1]))"]]; tasks.getRange("D5:D151").fillDown();
tasks.getRange("E5:E151").formulasR1C1 = [["=IF(RC[-2]=\"\",\"Undated\",TEXT(RC[-2],\"yyyy-mm\"))"]]; tasks.getRange("E5:E151").fillDown();
tasks.getRange("A4:N4").format = header; tasks.getRange("A4:N4").format.rowHeight = 38;
tasks.getRange("A5:N151").format.font = { size: 9, color: text };
tasks.getRange("A5:N151").format.borders = { preset: "insideHorizontal", style: "thin", color: "#D9E2F3" };
tasks.getRange("C5:C151").format.numberFormat = "m/d/yyyy"; tasks.getRange("K5:M151").format.numberFormat = "m/d/yyyy h:mm"; tasks.getRange("J5:J151").format.numberFormat = money;
tasks.getRange("A5:B151").format.wrapText = true; tasks.getRange("N5:N151").format.wrapText = true;
tasks.freezePanes.freezeRows(4); tasks.freezePanes.freezeColumns(2);
const taskWidths={A:18,B:38,C:14,D:11,E:12,F:22,G:20,H:12,I:22,J:14,K:19,L:19,M:19,N:42}; for(const [c,w] of Object.entries(taskWidths)) tasks.getRange(`${c}:${c}`).format.columnWidth=w;

// FINANCIAL LEDGER
title(ledger, "A1:I1", "ITEMIZED FINANCIAL LEDGER", "A2:I2", "Provisional dispute-support index. Reported Paid is not proof of payment; every entry requires confirmation against the authoritative financial system.");
ledger.getRange("A4:I4").values = [["Service Cycle", "Service Date", "Year", "Month", "Service Area", "Invoice(s)", "Classification", "Amount", "Source / Verification"]];
ledger.getRange("A5:I95").values = data.ledger.map(r => [r[0], r[1] ? new Date(r[1]) : null, null, null, r[2], r[3], r[4], r[5], r[6]]);
ledger.getRange("C5:C95").formulasR1C1 = [["=YEAR(RC[-1])"]]; ledger.getRange("C5:C95").fillDown();
ledger.getRange("D5:D95").formulasR1C1 = [["=TEXT(RC[-2],\"yyyy-mm\")"]]; ledger.getRange("D5:D95").fillDown();
ledger.getRange("A4:I4").format = header; ledger.getRange("A4:I4").format.rowHeight=38;
ledger.getRange("A5:I95").format.font={size:9,color:text}; ledger.getRange("A5:I95").format.borders={preset:"insideHorizontal",style:"thin",color:"#D9E2F3"};
ledger.getRange("B5:B95").format.numberFormat="m/d/yyyy"; ledger.getRange("H5:H95").format.numberFormat=money;
ledger.getRange("G5:G95").conditionalFormats.add("containsText",{text:"Denied",format:{fill:red,font:{bold:true,color:"#8A1C1C"}}});
ledger.getRange("G5:G95").conditionalFormats.add("containsText",{text:"Paid",format:{fill:green,font:{bold:true,color:"#38761D"}}});
ledger.freezePanes.freezeRows(4); const ledgerWidths={A:29,B:14,C:10,D:12,E:15,F:25,G:18,H:14,I:48}; for(const [c,w] of Object.entries(ledgerWidths)) ledger.getRange(`${c}:${c}`).format.columnWidth=w; ledger.getRange("A5:A95").format.wrapText=true; ledger.getRange("I5:I95").format.wrapText=true;

// YEARLY SUMMARY
title(yearly,"A1:I1","ASANA PERFORMANCE HISTORY - YEARLY SUMMARY","A2:I2","Formula-driven counts from the fictionalized 147-task dataset.");
yearly.getRange("A4:I4").values=[["Year","Total Tasks","Completed","Open","Failed Inspection","Passed / Compliant","Quality Review","No Status","Completion Rate"]];
yearly.getRange("A5:A11").values=data.yearly_labels.map(x=>[x]);
for(let r=5;r<=11;r++){
 yearly.getRange(`B${r}:I${r}`).formulas=[[`=COUNTIF('Asana Tasks'!$D$5:$D$151,A${r})`,`=COUNTIFS('Asana Tasks'!$D$5:$D$151,A${r},'Asana Tasks'!$H$5:$H$151,\"Yes\")`,`=B${r}-C${r}`,`=COUNTIFS('Asana Tasks'!$D$5:$D$151,A${r},'Asana Tasks'!$G$5:$G$151,\"Failed Inspection\")`,`=COUNTIFS('Asana Tasks'!$D$5:$D$151,A${r},'Asana Tasks'!$G$5:$G$151,\"Passed Inspection\")+COUNTIFS('Asana Tasks'!$D$5:$D$151,A${r},'Asana Tasks'!$G$5:$G$151,\"Compliant Received\")`,`=COUNTIFS('Asana Tasks'!$D$5:$D$151,A${r},'Asana Tasks'!$G$5:$G$151,\"Quality Review!\")`,`=COUNTIFS('Asana Tasks'!$D$5:$D$151,A${r},'Asana Tasks'!$G$5:$G$151,\"No status\")`,`=IFERROR(C${r}/B${r},0)`]];
}
yearly.getRange("A12:I12").values=[["TOTAL",null,null,null,null,null,null,null,null]]; yearly.getRange("B12:H12").formulas=[["=SUM(B5:B11)","=SUM(C5:C11)","=SUM(D5:D11)","=SUM(E5:E11)","=SUM(F5:F11)","=SUM(G5:G11)","=SUM(H5:H11)"]]; yearly.getRange("I12").formulas=[["=IFERROR(C12/B12,0)"]];
yearly.getRange("A4:I4").format=header; yearly.getRange("A12:I12").format={fill:pale,font:{bold:true,color:navy}}; yearly.getRange("I5:I12").format.numberFormat="0.0%"; yearly.getRange("A4:I12").format.borders={preset:"insideHorizontal",style:"thin",color:"#D9E2F3"};
for(const c of "ABCDEFGHI") yearly.getRange(`${c}:${c}`).format.columnWidth = c==="A"?14:17;

// MONTHLY SUMMARY
title(monthly,"A1:I1","ASANA PERFORMANCE HISTORY - MONTHLY SUMMARY","A2:I2","Monthly view based on each task's reporting date.");
monthly.getRange("A4:I4").values=[["Month","Total Tasks","Completed","Open","Failed Inspection","Passed / Compliant","Quality Review","No Status","Completion Rate"]];
monthly.getRange(`A5:A${4+data.monthly_labels.length}`).values=data.monthly_labels.map(x=>[x]);
for(let r=5;r<=4+data.monthly_labels.length;r++){
 monthly.getRange(`B${r}:I${r}`).formulas=[[`=COUNTIF('Asana Tasks'!$E$5:$E$151,A${r})`,`=COUNTIFS('Asana Tasks'!$E$5:$E$151,A${r},'Asana Tasks'!$H$5:$H$151,\"Yes\")`,`=B${r}-C${r}`,`=COUNTIFS('Asana Tasks'!$E$5:$E$151,A${r},'Asana Tasks'!$G$5:$G$151,\"Failed Inspection\")`,`=COUNTIFS('Asana Tasks'!$E$5:$E$151,A${r},'Asana Tasks'!$G$5:$G$151,\"Passed Inspection\")+COUNTIFS('Asana Tasks'!$E$5:$E$151,A${r},'Asana Tasks'!$G$5:$G$151,\"Compliant Received\")`,`=COUNTIFS('Asana Tasks'!$E$5:$E$151,A${r},'Asana Tasks'!$G$5:$G$151,\"Quality Review!\")`,`=COUNTIFS('Asana Tasks'!$E$5:$E$151,A${r},'Asana Tasks'!$G$5:$G$151,\"No status\")`,`=IFERROR(C${r}/B${r},0)`]];
}
monthly.getRange("A4:I4").format=header; monthly.getRange(`I5:I${4+data.monthly_labels.length}`).format.numberFormat="0.0%"; monthly.getRange(`A4:I${4+data.monthly_labels.length}`).format.borders={preset:"insideHorizontal",style:"thin",color:"#D9E2F3"}; monthly.freezePanes.freezeRows(4); for(const c of "ABCDEFGHI") monthly.getRange(`${c}:${c}`).format.columnWidth=c==="A"?14:17;

// FINANCIAL SUMMARY
title(financial,"A1:F1","FINANCIAL SUMMARY BY YEAR AND MONTH","A2:F2","Itemized amounts retained from the dispute analysis; financial-system verification remains outstanding.");
financial.getRange("A4:F4").values=[["Period","Reported Paid","Reported Denied","Itemized Total","Verified in Financial System?","Difference from Published Summary"]];
financial.getRange("A5:A16").values=data.financial_periods.map(x=>[x]);
for(let r=5;r<=16;r++){
 const annual=data.financial_periods[r-5].length===4;
 const periodCol=annual?"C":"D";
 financial.getRange(`B${r}:F${r}`).formulas=[[`=SUMIFS('Financial Ledger'!$H$5:$H$95,'Financial Ledger'!$${periodCol}$5:$${periodCol}$95,A${r},'Financial Ledger'!$G$5:$G$95,\"Reported Paid\")`,`=SUMIFS('Financial Ledger'!$H$5:$H$95,'Financial Ledger'!$${periodCol}$5:$${periodCol}$95,A${r},'Financial Ledger'!$G$5:$G$95,\"Reported Denied\")`,`=B${r}+C${r}`,`=\"No\"`,`=\"\"`]];
}
financial.getRange("A4:F4").format=header; financial.getRange("B5:D16").format.numberFormat=money; financial.getRange("A4:F16").format.borders={preset:"insideHorizontal",style:"thin",color:"#D9E2F3"}; const fw={A:14,B:18,C:18,D:18,E:26,F:32}; for(const [c,w] of Object.entries(fw)) financial.getRange(`${c}:${c}`).format.columnWidth=w;

// DATA QUALITY
title(quality,"A1:D1","DATA QUALITY AND RECONCILIATION FINDINGS","A2:D2","The audit separated reliable Asana operational evidence from financial claims requiring independent verification.");
quality.getRange("A4:D4").values=[["Issue","Observed Value","Impact","Required Resolution"]];
quality.getRange("A5:D15").values=[
 ["Live Asana task count",147,"Authoritative project snapshot population","Use the snapshot date when comparing later exports"],
 ["Tasks without a due / reporting date",21,"Cannot be assigned to a service month from Asana","Review attachments/comments and populate service date"],
 ["Tasks without a Status value",57,"Inspection outcome is not classifiable","Assign a documented status; do not infer"],
 ["Tasks with a nonblank Cost field",3,"Asana project is not a transaction ledger","Use the authoritative financial system for payment amounts"],
 ["Itemized Reported Paid total",139948,"Traceable to itemized summary rows; not proof of payment","Verify each invoice in the financial system"],
 ["Published paid summary total",152156,"Does not equal the itemized entries","Reconcile before reporting"],
 ["Paid-summary discrepancy",12208,"Unexplained overstatement in the older summary","Identify omitted or duplicated invoices"],
 ["Itemized Reported Denied total",23685,"Traceable to 11 itemized service-cycle rows","Verify denial and approval documentation"],
 ["Published denied summary total",25532,"Does not equal the itemized entries","Reconcile before reporting"],
 ["Denied-summary discrepancy",1847,"Unexplained overstatement in the older summary","Identify omitted or duplicated invoices"],
 ["Cost field value $194,904",194904,"Purchase-order amount, not an invoice-payment total","Exclude from invoice totals"],
 ];
quality.getRange("A4:D4").format=header; quality.getRange("A5:D15").format.borders={preset:"insideHorizontal",style:"thin",color:"#D9E2F3"}; quality.getRange("B9:B15").format.numberFormat=money; quality.getRange("A5:D15").format.wrapText=true; const qw={A:34,B:20,C:48,D:46}; for(const [c,w] of Object.entries(qw)) quality.getRange(`${c}:${c}`).format.columnWidth=w;

// DASHBOARD
title(dashboard,"A1:H1","ASANA VENDOR PERFORMANCE AND INVOICE AUDIT","A2:H2","Management-requested payment-dispute analysis | Fictional identities | Original analytical figures preserved");
dashboard.getRange("A4:B4").values=[["ASANA RECORDS",""]]; dashboard.getRange("A4:B4").format=section;
dashboard.getRange("A5:A9").values=[["Total Tasks"],["Completed"],["Open"],["Failed Inspection"],["Undated"]];
dashboard.getRange("B5:B9").formulas=[["=COUNTA('Asana Tasks'!$A$5:$A$151)"],["=COUNTIF('Asana Tasks'!$H$5:$H$151,\"Yes\")"],["=COUNTIF('Asana Tasks'!$H$5:$H$151,\"No\")"],["=COUNTIF('Asana Tasks'!$G$5:$G$151,\"Failed Inspection\")"],["=COUNTIF('Asana Tasks'!$D$5:$D$151,\"Undated\")"]];
dashboard.getRange("D4:E4").values=[["ITEMIZED FINANCIALS",""]]; dashboard.getRange("D4:E4").format=section;
dashboard.getRange("D5:D10").values=[["Reported Paid"],["Denied - Work Not Performed"],["Itemized Total"],["Vendor Unpaid Claim"],["Payable Unpaid Amount Found"],["Financial-System Verified"]];
dashboard.getRange("E5:E7").formulas=[["=SUMIF('Financial Ledger'!$G$5:$G$95,\"Reported Paid\",'Financial Ledger'!$H$5:$H$95)"],["=SUMIF('Financial Ledger'!$G$5:$G$95,\"Reported Denied\",'Financial Ledger'!$H$5:$H$95)"],["=E5+E6"]];
dashboard.getRange("E8:E10").values=[["> $40,000"],[0],["No"]]; dashboard.getRange("E5:E7").format.numberFormat=money; dashboard.getRange("E9").format.numberFormat=money;
dashboard.getRange("A5:B9").format={fill:pale,borders:{preset:"insideHorizontal",style:"thin",color:"#D9E2F3"}}; dashboard.getRange("D5:E10").format={fill:pale,borders:{preset:"insideHorizontal",style:"thin",color:"#D9E2F3"}};
dashboard.getRange("A11:H11").merge(); dashboard.getRange("A11").values=[["MANAGEMENT FINDING"]]; dashboard.getRange("A11:H11").format=section;
dashboard.getRange("A12:H13").merge(); dashboard.getRange("A12").values=[["The vendor asserted that more than $40,000 remained unpaid. The itemized review found $0 payable for completed work remained unpaid. The $23,685 denied amount corresponded to work that was not performed or did not satisfy documented requirements; it was therefore not an unpaid obligation. The analysis gave management a detailed, traceable basis for the organization's position. Asana evidence still required reconciliation to the authoritative financial system before final payment certification."]]; dashboard.getRange("A12:H13").format={fill:yellow,font:{bold:true,color:"#7F6000"},wrapText:true,verticalAlignment:"center"};
dashboard.getRange("A15:H15").merge(); dashboard.getRange("A15").values=[["PUBLIC PORTFOLIO NOTICE: Vendor, employee, location, task, invoice, link, and municipal identifiers are fictional. Analytical counts and monetary figures are retained from the management-requested audit."]]; dashboard.getRange("A15:H15").format={font:{italic:true,color:"#666666"},wrapText:true};
const dw={A:26,B:14,C:4,D:28,E:18,F:8,G:8,H:8}; for(const [c,w] of Object.entries(dw)) dashboard.getRange(`${c}:${c}`).format.columnWidth=w; dashboard.getRange("A12:H13").format.rowHeight=40; dashboard.getRange("A15:H15").format.rowHeight=34;

const out=await SpreadsheetFile.exportXlsx(wb);
const path=`${repo}/Asana_Vendor_Performance_and_Invoice_Audit_Fictionalized.xlsx`;
await out.save(path);
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:200},summary:"formula error scan"}); console.log(errors.ndjson);
for(const [sheetName,range,name] of [["Dashboard","A1:H15","dashboard_preview.png"],["Data Quality","A1:D15","data_quality_preview.png"]]){const img=await wb.render({sheetName,range,scale:1.2,format:"png"}); await fs.writeFile(`${repo}/${name}`,new Uint8Array(await img.arrayBuffer()));}
console.log(path);
