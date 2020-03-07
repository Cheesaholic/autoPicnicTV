function main() {
  var spreadsheet = '';
  var slideShift1 = '';
  var slideShift2 = '';
  var slideShift3 = '';
  var slideBHV = '';
  
  var ids = [spreadsheet, slideShift1, slideShift2, slideShift3, slideBHV]
  getSheet(ids);
  autoBHV(ids);
}

function getSheet(ids) {
  removeTables(ids);
  var sheetActive = SpreadsheetApp.openById(ids[0]);
  var sheet = sheetActive.getSheetByName(getDate(0));
  if (sheet) {
    var runners = getList(sheet, 'N:N');
    var vtijden = getList(sheet, 'I:I');
    
    var deliveryZones = getDeliveryZones(ids[0]);
    
    createTables(ids, 1, runners[0], vtijden[0], deliveryZones);
    createTables(ids, 2, runners[1], vtijden[1], deliveryZones);
    createTables(ids, 3, runners[2], vtijden[2], deliveryZones);
    
  }
  else {
    SlidesApp.openById(ids[1]).getSlideById('p').getShapes()[0].getText().setText('Start day to show inlaadvolgorde');
    SlidesApp.openById(ids[2]).getSlideById('p').getShapes()[0].getText().setText('Start day to show inlaadvolgorde');
    SlidesApp.openById(ids[3]).getSlideById('p').getShapes()[0].getText().setText('Start day to show inlaadvolgorde');
    removeTables(ids);
  }
}

function autoBHV(ids) {
  var sheetActive = SpreadsheetApp.openById(ids[0]);
  var sheet = sheetActive.getSheetByName('Cockpit');
  if (!sheet.getRange('C1').isBlank()) {
    var BHV = sheet.getRange('C1').getValue();
    BHV = BHV.replace(/name/gi, 'Name(Title)');
    var slidesPage = SlidesApp.openById(ids[4]).getSlideById('p');
    if(slidesPage.getShapes()[0].getText().asString() != BHV) {
      slidesPage.getShapes()[0].getText().setText(BHV);
    }
  }
  else{
    SlidesApp.openById(ids[4]).getSlideById('p').getShapes()[0].getText().setText('No R+/HM present, please start day');
  }
}

function removeTables(ids) {
  var shift1 = SlidesApp.openById(String(ids[1])).getSlideById('p');
  var shift2 = SlidesApp.openById(ids[2]).getSlideById('p');
  var shift3 = SlidesApp.openById(ids[3]).getSlideById('p');
  
  var shifts = [shift1, shift2, shift3]
  
  shifts.forEach(remfuncs);
}

function remfuncs(slide) {
  slide.getTables().forEach(remtable);
  slide.getShapes()[1].getText().setText(' ');
  slide.getImages().forEach(remtable);
}

function createTables(ids, shift, values, vtijden, deliveryZones) {
  if(deliveryZones[shift-1] == 'deliveryzone1') {
    var color = '#0000ff';
  } else if(deliveryZones[shift-1] == 'deliveryzone2') {
    var color = '#6aa84f';
  } else if(deliveryZones[shift-1] == 'deliveryzone3') {
    var color = '#e69318';
  } else {
    var color = '#9b1c31'; }
  if(shift == 1) {
    var slidesPage = SlidesApp.openById(ids[1]).getSlideById('p');
    slidesPage.getShapes()[0].getText().setText('Inlaadvolgorde ' + getDate(1) + ' - Shift ' + String(shift) + ' - ');
    slidesPage.getShapes()[1].getText().setText(deliveryZones[shift-1]).getTextStyle().setForegroundColor(color);
    
  } else if(shift == 2) {
    var slidesPage = SlidesApp.openById(ids[2]).getSlideById('p');
    slidesPage.getShapes()[0].getText().setText('Inlaadvolgorde ' + getDate(1) + ' - Shift ' + String(shift) + ' - '); 
    slidesPage.getShapes()[1].getText().setText(deliveryZones[shift-1]).getTextStyle().setForegroundColor(color);
  } else if(shift == 3) {
    var slidesPage = SlidesApp.openById(ids[3]).getSlideById('p');
    slidesPage.getShapes()[0].getText().setText('Inlaadvolgorde ' + getDate(1) + ' - Shift ' + String(shift) + ' - '); 
    slidesPage.getShapes()[1].getText().setText(deliveryZones[shift-1]).getTextStyle().setForegroundColor(color);
  }
//  slidesPage.insertImage(deliveryZones[shift+2], 530, 5, 40, 40);
  var fontsize = 12;
  var columns = 2;
  var rows = values.length;
  var for1 = 8;
  if(values.length < 8) {
    for1 = values.length;
    }
  var table1 = slidesPage.insertTable(for1+1, columns,25, 50, 190, 10);
  table1.getCell(0, 1).getText().setText('Naam').getTextStyle().setBold(true).setFontSize(12);
  table1.getCell(0, 0).getText().setText('Vertrektijd').getTextStyle().setBold(true).setFontSize(12);
  for (var r = 0; r < for1; r++) {
    p = r + 1;
    if(values[r] == 'NEED RUNNER') {
      table1.getCell(p, 1).getText().setText(values[r]).getTextStyle().setFontSize(fontsize).setForegroundColor('#ffffff').setBold(true);
      table1.getCell(p, 1).getFill().setSolidFill('#e40611');
      table1.getCell(p, 0).getText().setText(vtijden[r]).getTextStyle().setFontSize(fontsize).setForegroundColor('#ffffff').setBold(true);
      table1.getCell(p, 0).getFill().setSolidFill('#e40611');
    } else {
      table1.getCell(p, 1).getText().setText(values[r]).getTextStyle().setFontSize(fontsize);
      table1.getCell(p, 0).getText().setText(vtijden[r]).getTextStyle().setFontSize(fontsize)
    }
  }
  if(values.length > 8) {
    var for2 = 16
    if(values.length < 16) {
      for2 = values.length;
    }
    var table2 = slidesPage.insertTable(for2-for1+1, columns,265, 50, 190, 10);
    table2.getCell(0, 1).getText().setText('Naam').getTextStyle().setBold(true).setFontSize(12);
    table2.getCell(0, 0).getText().setText('Vertrektijd').getTextStyle().setBold(true).setFontSize(12);
    var p2 = 1
    for (var r = 8; r < for2; r++) {
      if(values[r] == 'NEED RUNNER') {
      table2.getCell(p2, 1).getText().setText(values[r]).getTextStyle().setFontSize(fontsize).setForegroundColor('#ffffff').setBold(true);
      table2.getCell(p2, 1).getFill().setSolidFill('#e40611');
      table2.getCell(p2, 0).getText().setText(vtijden[r]).getTextStyle().setFontSize(fontsize).setForegroundColor('#ffffff').setBold(true);
      table2.getCell(p2, 0).getFill().setSolidFill('#e40611');
    } else {
      table2.getCell(p2, 1).getText().setText(values[r]).getTextStyle().setFontSize(fontsize);
      table2.getCell(p2, 0).getText().setText(vtijden[r]).getTextStyle().setFontSize(fontsize)
    }
      p2++;
    }
    if(values.length > 16) {
    var table3 = slidesPage.insertTable(values.length-15, columns,500, 50, 190, 10);
    table3.getCell(0, 1).getText().setText('Naam').getTextStyle().setBold(true).setFontSize(12);
    table3.getCell(0, 0).getText().setText('Vertrektijd').getTextStyle().setBold(true).setFontSize(12);
    var p3 = 1;
    for (var r = 16; r < values.length; r++) {
      if(values[r] == 'NEED RUNNER') {
      table3.getCell(p3, 1).getText().setText(values[r]).getTextStyle().setFontSize(fontsize).setForegroundColor('#ffffff').setBold(true);
      table3.getCell(p3, 1).getFill().setSolidFill('#e40611');
      table3.getCell(p3, 0).getText().setText(vtijden[r]).getTextStyle().setFontSize(fontsize).setForegroundColor('#ffffff').setBold(true);
      table3.getCell(p3, 0).getFill().setSolidFill('#e40611');
    } else {
      table3.getCell(p3, 1).getText().setText(values[r]).getTextStyle().setFontSize(fontsize);
      table3.getCell(p3, 0).getText().setText(vtijden[r]).getTextStyle().setFontSize(fontsize)
    }
      p3++;
    }
    }
  }
  }
  
  
  
function remtable(table) {
  table.remove();
}

function getList(sheet, range) {
  var column = sheet.getRange(range);
  var values = column.getDisplayValues();
  var shift1Run = filterArray(values, '');
  shift1Run.splice(0,1);
  for(var indexRun2 in shift1Run){
    if((shift1Run[indexRun2][0]=='Runner') || (shift1Run[indexRun2][0]=='Vertrek')){break}
  }
  var shift2Run = shift1Run.splice(indexRun2);
  shift2Run.splice(0,1);
  
  for(var indexRun3 in shift2Run){
    if((shift2Run[indexRun3][0]=='Runner') || (shift2Run[indexRun3][0]=='Vertrek')){break}
  }
  var shift3Run = shift2Run.splice(indexRun3);
  shift3Run.splice(0,1);

  return [shift1Run, shift2Run, shift3Run]
}

function dateTimeToString(array) {
  array.forEach(function(part, index, theArray) {
    theArray[index] = part.getHours() + ':' + part.getMinutes()
});
}

function getDeliveryZone(zoneString) {
  zone = '';
  zoneImg= '';
  if(String(zoneString).indexOf('deliveryzone1') != -1) {
    zone = 'Zuid / Geldrop';
    zoneImg = 'https://cdn.onlinewebfonts.com/svg/img_15731.png';
  } else if(String(zoneString).indexOf('deliveryzone2') != -1) {
      zone = 'Meerhoven';
    zoneImg = 'https://cdn.onlinewebfonts.com/svg/img_15726.png';
    } else if(String(zoneString).indexOf('deliveryzone3') != -1) {
        zone = 'Noord';
      zoneImg = 'https://cdn.onlinewebfonts.com/svg/img_15726.png';
      }
  return [zone, zoneImg];
}

function getDeliveryZones(spreadsheet) {
  var sheetActive = SpreadsheetApp.openById(spreadsheet);
  var sheet = sheetActive.getSheetByName(getDate(0));
  var zoneStringShift1 = sheet.getRange(6,18).getDisplayValue();
  for (var rShift2 = 7; rShift2 < 300; rShift2++) {
    if(sheet.getRange(rShift2,18).getDisplayValue() == 'Area') {
      rShift2++;
      var zoneStringShift2 = sheet.getRange(rShift2,18).getDisplayValue();
      rShift2++;
      break;
    }}
  for (var rShift3 = rShift2; rShift3 < 300; rShift3++) {
    if(sheet.getRange(rShift3,18).getDisplayValue() == 'Area') {
      rShift3++;
      var zoneStringShift3 = sheet.getRange(rShift3,18).getDisplayValue();
      rShift3++;
      break;
    }
  }
  var zoneShift1 = getDeliveryZone(zoneStringShift1);
  var zoneShift2 = getDeliveryZone(zoneStringShift2);
  var zoneShift3 = getDeliveryZone(zoneStringShift3);
  
  zones = [zoneShift1[0], zoneShift2[0], zoneShift3[0], zoneShift1[1], zoneShift2[1], zoneShift3[1]];
  return zones

}


function getDate(num) {
  if(num == 0) {
    return Utilities.formatDate(new Date(), "GMT+1", "dd_MM");
  }
  else{
    return Utilities.formatDate(new Date(), "GMT+1", "dd-MM");
  }
}

function filterArray(arr, str) {
  return filtered = arr.filter(function (el) {
    return el != str;
    });
}