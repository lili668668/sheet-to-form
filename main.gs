function onOpen () {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('製作表單')
    .addItem('製作新表單', 'makeForm')
    .addToUi();
}

function makeForm () {
  var spreadsheet = SpreadsheetApp.getActive();
  var name = spreadsheet.getName();
  var form = FormApp.create(name);
  var sheets = spreadsheet.getSheets();
  var titleSetted = false;
  var sectionMap = {};
  var sections = [];
  for (var cnt = 0;cnt < sheets.length;cnt++) {
    var sheet = sheets[cnt];
    var items = [];
    if (sheet.getName() == 'setting'){
      form.setConfirmationMessage(getValue(sheet, 'B1'))
        .setCollectEmail(getValue(sheet, 'B2'))
        .setAllowResponseEdits(getValue(sheet, 'B3'))
        .setPublishingSummary(getValue(sheet, 'B4'))
        .setProgressBar(getValue(sheet, 'B5'))
        .setShuffleQuestions(getValue(sheet, 'B6'))
        .setShowLinkToRespondAgain(getValue(sheet, 'B7'))
      sections.push(items)
      continue;
    }

    if (sheet.getName().split('Ignore').length >= 2) {
      sections.push(items)
      continue;
    }

    if (!titleSetted) {
      form.setTitle(getValue(sheet, 'B1'))
        .setDescription(getValue(sheet, 'B2'));
      sectionMap[sheet.getName()] = FormApp.PageNavigationType.RESTART;
      titleSetted = true;
    } else {
      var pageBreakItem = form.addPageBreakItem()
        .setTitle(getValue(sheet, 'B1'))
        .setHelpText(getValue(sheet, 'B2'))
      sections[cnt - 1].push(pageBreakItem);
      sectionMap[sheet.getName()] = pageBreakItem;
    }

    var list = sheet.getRange('A5:A').getValues();
    for (var cnt2 = 0;cnt2 < list.length;cnt2++) {
      var type = list[cnt2][0]
      var rowNum = 5 + cnt2;
      switch (type) {
        case '文字描述':
          var sectionHeaderItem = form.addSectionHeaderItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          items.push(sectionHeaderItem);
          break;
        case '圖片':
          var imageItem = form.addImageItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue())
            .setImage(UrlFetchApp.fetch(sheet.getRange(rowNum, 5).getValue()));
          items.push(imageItem);
          break;
        case '影片':
          var videoItem = form.addVideoItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue())
            .setVideoUrl(sheet.getRange(rowNum, 5).getValue());
          items.push(videoItem);
          break;
        case '選擇題':
          var multipleChoiceItem = form.addMultipleChoiceItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            multipleChoiceItem.setRequired(true);
          }
          items.push(multipleChoiceItem);
          break;
        case '簡答題':
          var textItem = form.addTextItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            textItem.setRequired(true);
          }
          textItem.setValidation(createTextValidation(sheet.getRange(rowNum, 10).getValue(), sheet.getRange(rowNum, 6).getValue(), sheet.getRange(rowNum, 9).getValue()));
          items.push(textItem);
          break;
        case '段落':
          var paragraphTextItem = form.addParagraphTextItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            paragraphTextItem.setRequired(true);
          }
          paragraphTextItem.setValidation(createParagraphTextValidation(sheet.getRange(rowNum, 10).getValue(), sheet.getRange(rowNum, 7).getValue(), sheet.getRange(rowNum, 9).getValue()));
          items.push(paragraphTextItem);
          break;
        case '核取方塊':
          var checkboxItem = form.addCheckboxItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            checkboxItem.setRequired(true);
          }
          checkboxItem.setValidation(createCheckboxValidation(sheet.getRange(rowNum, 10).getValue(), sheet.getRange(rowNum, 8).getValue(), sheet.getRange(rowNum, 9).getValue()));
          var choices = sheet.getRange(rowNum, 5).getValue().split('\n');
          if (choices[choices.length - 1] == '其他') {
            checkboxItem.showOtherOption(true);
            checkboxItem.setChoiceValues(choices.slice(0, choices.length - 1));
          } else {
            checkboxItem.setChoiceValues(choices)
          }
          items.push(checkboxItem);
          break;
        case '下拉式選單':
          var listItem = form.addListItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            listItem.setRequired(true);
          }
          items.push(listItem);
          break;
        case '線性刻度':
          var scaleItem = form.addScaleItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            scaleItem.setRequired(true);
          }
          var options = sheet.getRange(rowNum, 5).getValue().split('\n');
          var small = options[0].split(',');
          var big = options[1].split(',');
          scaleItem
            .setBounds(small[0], big[0])
            .setLabels(small[1], big[1]);
          items.push(scaleItem);
          break;
        case '單選方格':
          var gridItem = form.addGridItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            gridItem.setRequired(true);
          }
          var options = sheet.getRange(rowNum, 5).getValue().split('\n');
          var columns = options[0].split(',');
          var rows = options[1].split(',');
          if (options[2] == '每一欄僅限一則回應') {
            gridItem.setValidation(FormApp.createGridValidation().requireLimitOneResponsePerColumn().build());
          }
          gridItem
            .setColumns(columns)
            .setRows(rows);
          items.push(gridItem);
          break;
        case '核取方塊格':
          var checkboxGridItem = form.addCheckboxGridItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            checkboxGridItem.setRequired(true);
          }
          var options = sheet.getRange(rowNum, 5).getValue().split('\n');
          var columns = options[0].split(',');
          var rows = options[1].split(',');
          if (options[2] == '每一欄僅限一則回應') {
            checkboxGridItem.setValidation(FormApp.createCheckboxGridValidation().requireLimitOneResponsePerColumn().build());
          }
          checkboxGridItem
            .setColumns(columns)
            .setRows(rows);
          items.push(checkboxGridItem);
          break;
        case '日期':
          var dateItem = form.addDateItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            dateItem.setRequired(true);
          }
          if (sheet.getRange(rowNum, 5).getValue() == '加入年份') {
            dateItem.setIncludesYear(true);
          }
          items.push(dateItem);
          break;
        case '日期與時間':
          var dateTimeItem = form.addDateTimeItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            dateTimeItem.setRequired(true);
          }
          if (sheet.getRange(rowNum, 5).getValue() == '加入年份') {
            dateTimeItem.setIncludesYear(true);
          }
          items.push(dateTimeItem);
          break;
        case '時間':
          var timeItem = form.addTimeItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            timeItem.setRequired(true);
          }
          items.push(timeItem);
          break;
        case '持續時間（時間長度）':
          var dTimeItem = form.addDurationItem()
            .setTitle(sheet.getRange(rowNum, 2).getValue())
            .setHelpText(sheet.getRange(rowNum, 3).getValue());
          if (sheet.getRange(rowNum, 4).getValue()) {
            dTimeItem.setRequired(true);
          }
          items.push(dTimeItem);
          break;
        default:
          items.push(null);
      }
    }

    sections.push(items);
  }

  for (var cnt = 0;cnt < sheets.length;cnt++) {
    var sheet = sheets[cnt]
    if (sheet.getName() == 'setting' || sheet.getName().split('Ignore').length >= 2) continue;
    var pageBreakItem = sections[cnt][sections[cnt].length - 1];
    if (pageBreakItem != null && pageBreakItem.getType() == FormApp.ItemType.PAGE_BREAK) {
      pageBreakItem.setGoToPage(getPageBreakItem(getValue(sheet, 'B3'), sectionMap));
    }
  }

  for (var cnt = 0;cnt < sheets.length;cnt++) {
    var sheet = sheets[cnt]
    if (sheet.getName() == 'setting' || sheet.getName().split('Ignore').length >= 2) continue;
    for (var cnt2 = 0;cnt2 < list.length;cnt2++) {
      var type = list[cnt2][0]
      var rowNum = 5 + cnt2;
      switch (type) {
        case '選擇題':
        case '下拉式選單':
          var choices = []
          var options = sheet.getRange(rowNum, 5).getValue().split('\n')
          for (var cnt3 = 0;cnt3 < options.length;cnt3++) {
            var items = options[cnt3].split('->')
            if (items[0].trim() == '' || items[0].trim() == null) continue;
            if (items[1] == '' || items[1] == undefined || items[1] == null) {
              choices.push(sections[cnt][cnt2].createChoice(items[0].trim()))
            } else {
              choices.push(sections[cnt][cnt2].createChoice(items[0].trim(), getPageBreakItem(items[1].trim(), sectionMap)));
            }
          }
          if (sections[cnt][cnt2] != null) sections[cnt][cnt2].setChoices(choices);
          break;
        default:
      }
    }
  }

  SpreadsheetApp.getUi().alert('表單：' + form.getEditUrl());
  settingSheet = spreadsheet.getSheetByName('setting');
  settingSheet.getRange(settingSheet.getLastRow() + 1, 1).setValue('（自動產生）產生過的表單：' + form.getEditUrl());
}

function getPageBreakItem (key, sectionMap) {
  switch(key) {
    case '提交':
      return FormApp.PageNavigationType.SUBMIT;
      break;
    case '前往下一個區段':
      return FormApp.PageNavigationType.CONTINUE;
      break;
    default:
      return sectionMap[key];
  }
}

function createTextValidation (text, condition, option) {
  switch (condition) {
    case '數字大於':
      return FormApp.createTextValidation()
        .requireNumberGreaterThan(option)
        .setHelpText(text)
        .build();
    case '數字大於或等於':
      return FormApp.createTextValidation()
        .requireNumberGreaterThanOrEqualTo(option)
        .setHelpText(text)
        .build();
    case '數字小於':
      return FormApp.createTextValidation()
        .requireNumberLessThan(option)
        .setHelpText(text)
        .build();
    case '數字小於或等於':
      return FormApp.createTextValidation()
        .requireNumberLessThanOrEqualTo(option)
        .setHelpText(text)
        .build();
    case '數字等於':
      return FormApp.createTextValidation()
        .requireNumberEqualTo(option)
        .setHelpText(text)
        .build();
    case '數字不等於':
      return FormApp.createTextValidation()
        .requireNumberNotEqualTo(option)
        .setHelpText(text)
        .build();
    case '數字介於（需要輸入兩個數字，用半形逗號區隔）':
      var nums = option.split(',');
      return FormApp.createTextValidation()
        .requireNumberBetween(nums[0].trim(), nums[1].trim())
        .setHelpText(text)
        .build();
    case '數字非介於（需要輸入兩個數字，用半形逗號區隔）':
      var nums = option.split(',');
      return FormApp.createTextValidation()
        .requireNumberNotBetween(nums[0].trim(), nums[1].trim())
        .setHelpText(text)
        .build();
    case '數字是數字（含浮點數）':
      return FormApp.createTextValidation()
        .requireNumber()
        .setHelpText(text)
        .build();
    case '數字是整數':
      return FormApp.createTextValidation()
        .requireWholeNumber()
        .setHelpText(text)
        .build();
    case '文字包含':
      return FormApp.createTextValidation()
        .requireTextContainsPattern(option)
        .setHelpText(text)
        .build();
    case '文字不包含':
      return FormApp.createTextValidation()
        .requireTextDoesNotContainPattern(option)
        .setHelpText(text)
        .build();
    case '文字是電子郵件地址':
      return FormApp.createTextValidation()
        .requireTextIsEmail()
        .setHelpText(text)
        .build();
    case '文字是網址':
      return FormApp.createTextValidation()
        .requireTextIsUrl()
        .setHelpText(text)
        .build();
    case '長度的最大字元數':
      return FormApp.createTextValidation()
        .requireTextLengthLessThanOrEqualTo(option)
        .setHelpText(text)
        .build();
    case '長度的最小字元數':
      return FormApp.createTextValidation()
        .requireTextLengthGreaterThanOrEqualTo(option)
        .setHelpText(text)
        .build();
    default:
      return FormApp.createTextValidation().build();
  }
}

function createParagraphTextValidation (text, condition, option) {
  switch (condition) {
    case '長度的最大字元數':
      return FormApp.createParagraphTextValidation()
        .requireTextLengthLessThanOrEqualTo(option)
        .setHelpText(text)
        .build();
    case '長度最小字元數':
      return FormApp.createParagraphTextValidation()
        .requireTextLengthGreaterThanOrEqualTo(option)
        .setHelpText(text)
        .build();
    default:
      return FormApp.createParagraphTextValidation().build();
  }
}

function createCheckboxValidation (text, condition, option) {
  switch (condition) {
    case '選取至少':
      return FormApp.createCheckboxValidation()
        .requireSelectAtLeast(option)
        .setHelpText(text)
        .build();
    case '選取至多':
      return FormApp.createCheckboxValidation()
        .requireSelectAtMost(option)
        .setHelpText(text)
        .build();
    case '選取剛好':
      return FormApp.createCheckboxValidation()
        .requireSelectExactly(option)
        .setHelpText(text)
        .build();
    default:
      return FormApp.createCheckboxValidation().build();
  }
}

function getValue(sheet, cell) {
  var value = sheet.getRange(cell).getValue();
  return value;
}

function consolelog (value) {
  SpreadsheetApp.getUi()
    .alert(value);
}
