const GRAPHIC_SHEET="Graphic";

function createGraphicSheet(){
  const ss=getSpreadsheet_(),period=getCurrentPeriod_();let sh=ss.getSheetByName(GRAPHIC_SHEET);if(!sh)sh=ss.insertSheet(GRAPHIC_SHEET);
  const summary={FitFood:{count:0,money:0},ServiceFood:{count:0,money:0},Hool:{count:0,money:0},OglooniiTsai:{count:0,money:0}},daily={};
  const configs=[{sheet:FIT_SHEET,key:"FitFood",price:15000},{sheet:SERVICE_SHEET,key:"ServiceFood",service:true},{sheet:HOOL_SHEET,key:"Hool",price:15000},{sheet:BREAKFAST_SHEET,key:"OglooniiTsai",price:12000}];
  configs.forEach(cfg=>{const src=ss.getSheetByName(cfg.sheet);if(!src)return;const data=src.getDataRange().getValues();if(data.length<2)return;const totalCol=data[0].indexOf("Нийт");if(totalCol<0)return;
    for(let c=4;c<totalCol;c++){if(!isDateInCurrentPeriod_(data[0][c]))continue;const dateKey=graphicDateKey_(data[0][c],period);if(!daily[dateKey])daily[dateKey]={FitFood:0,ServiceFood:0,Hool:0,OglooniiTsai:0};
      for(let r=1;r<data.length;r++){if(String(data[r][0]).trim()==="Нийт")continue;const n=Number(data[r][c]);if(!Number.isFinite(n)||n<=0)continue;if(cfg.service){summary[cfg.key].count+=1;summary[cfg.key].money+=n;daily[dateKey][cfg.key]+=1;}else{summary[cfg.key].count+=n;summary[cfg.key].money+=n*cfg.price;daily[dateKey][cfg.key]+=n;}}
    }
  });
  const totalCount=summary.FitFood.count+summary.ServiceFood.count+summary.Hool.count+summary.OglooniiTsai.count,totalFood=summary.FitFood.count+summary.Hool.count+summary.OglooniiTsai.count,totalMoney=Object.keys(summary).reduce((s,k)=>s+summary[k].money,0);
  sh.getCharts().forEach(c=>sh.removeChart(c));sh.clear();
  sh.getRange("A1:H1").merge().setValue("АКУМА • ХООЛНЫ НЭГДСЭН ГРАФИК").setFontSize(18).setFontWeight("bold").setFontColor("white").setBackground("#17233c").setHorizontalAlignment("center");
  sh.getRange("A2:H2").merge().setValue("Тооцооны хугацаа: "+period.label).setFontWeight("bold").setFontColor("#1d4ed8").setBackground("#eff6ff").setHorizontalAlignment("center");
  sh.getRange("A4:D4").setValues([["Хоолны төрөл","Бүртгэлийн тоо","Нэгж үнэ","Мөнгөн дүн"]]);
  const summaryRows=[["FitFood",summary.FitFood.count,15000,summary.FitFood.money],["ServiceFood",summary.ServiceFood.count,"Үнийн дүнгээр",summary.ServiceFood.money],["Хоол",summary.Hool.count,15000,summary.Hool.money],["Өглөөний цай",summary.OglooniiTsai.count,12000,summary.OglooniiTsai.money],["Нийт",totalCount,"Нийт хоол: "+totalFood,totalMoney]];
  sh.getRange(5,1,summaryRows.length,4).setValues(summaryRows);sh.getRange("C5:C8").setNumberFormat("#,##0₮");sh.getRange("D5:D9").setNumberFormat("#,##0₮");
  const dates=Object.keys(daily).sort(),dailyRows=dates.map(d=>[d,daily[d].FitFood,daily[d].ServiceFood,daily[d].Hool,daily[d].OglooniiTsai,daily[d].FitFood+daily[d].ServiceFood+daily[d].Hool+daily[d].OglooniiTsai]);
  sh.getRange("A12:F12").setValues([["Огноо","FitFood","ServiceFood","Хоол","Өглөөний цай","Нийт бүртгэл"]]);if(dailyRows.length)sh.getRange(13,1,dailyRows.length,6).setValues(dailyRows);
  [sh.getRange("A4:D4"),sh.getRange("A12:F12")].forEach(r=>r.setFontWeight("bold").setFontColor("white").setBackground("#2563eb").setHorizontalAlignment("center"));sh.getRange("A9:D9").setFontWeight("bold").setBackground("#dbeafe");sh.setFrozenRows(2);sh.setColumnWidth(1,150);sh.setColumnWidths(2,5,125);sh.setColumnWidth(7,30);sh.setColumnWidth(8,30);
  const pie=sh.newChart().setChartType(Charts.ChartType.PIE).addRange(sh.getRange("A4:B8")).setPosition(4,6,0,0).setOption("title","Хоолны төрлүүдийн бүртгэлийн хувь").setOption("pieHole",0.42).setOption("legend",{position:"right"}).build();
  const money=sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(sh.getRange("A4:A8")).addRange(sh.getRange("D4:D8")).setPosition(20,6,0,0).setOption("title","Хоолны төрлүүдийн мөнгөн дүн").setOption("legend",{position:"none"}).setOption("colors",["#2563eb"]).build();sh.insertChart(pie);sh.insertChart(money);
  if(dailyRows.length){const trend=sh.newChart().setChartType(Charts.ChartType.LINE).addRange(sh.getRange(12,1,dailyRows.length+1,6)).setPosition(20,1,0,0).setOption("title","Өдөр тутмын бүртгэлийн хөдөлгөөн").setOption("curveType","function").setOption("pointSize",4).build();sh.insertChart(trend);}
  return{success:true,period:period.label,totalFood:totalFood,totalMoney:totalMoney};
}

function graphicDateKey_(value,period){const text=String(value||"").trim();if(/^\d{4}-\d{1,2}-\d{1,2}$/.test(text))return text;const p=normalizeDate(text).split("/").map(Number);let year=period.start.getFullYear();if(p[0]<period.start.getMonth()+1)year=period.end.getFullYear();return year+"-"+String(p[0]).padStart(2,"0")+"-"+String(p[1]).padStart(2,"0");}
