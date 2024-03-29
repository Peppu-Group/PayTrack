const SELLING_PRICE = 'Selling Price';
const PAYMENT_METHOD = 'Payment Method';
const BUYING_PRICE = 'Buying Price';
const EXP_CATEGORY = 'Choose Expense Category';
const INFLOW_CATEGORY = 'Choose Inflow Category';
const LOAN_CATEGORY = 'Choose Liability Type';

var newSheetSection = CardService.newCardSection();
var inputSheetSection = CardService.newCardSection();
var buttonSheetSection = CardService.newCardSection();
var invoiceSection = CardService.newCardSection();
var navigationSection = CardService.newCardSection();
var sendInvoiceSection = CardService.newCardSection();
var selectInvoiceSection = CardService.newCardSection();
var transactionSection = CardService.newCardSection();

const INPUT_MAP = [
  { text: 'Bank', val: 'Bank' },
  { text: 'Cash', val: 'Cash' },
  { text: 'Loan', val: 'Loan' },
  { text: 'Credit card', val: 'Credit card' },
  { text: 'Sales', val: 'Sales' },
  { text: 'Services', val: 'Services' },
]

const OUT_MAP = [
  { text: 'Rent', val: 'Rent' },
  { text: 'Utility', val: 'Utility' },
  { text: 'Office Supplies', val: 'Office Supplies' },
  { text: 'Advertising', val: 'Advertising' },
  { text: 'Entertainment', val: 'Entertainment' },
  { text: 'Other Expenses', val: 'Other Expenses' },
]

const PAYMENT_MAP = [
  { text: 'Bank', val: 'Bank' },
  { text: 'Cash', val: 'Cash' }
]

const LIABILITY_MAP = [
  { text: 'Loan', val: 'Loan' },
  { text: 'Credit card', val: 'Credit card' }
]

function invoice_card() {
  /*The reason we're writing this out, is so that we can call the invcard build function
    anywhere. */
  // create invoice
  var contactName = CardService.newTextInput()
    .setFieldName(`Contact Name`)
    .setTitle(`Receiver's name`);

  var clientName = CardService.newTextInput()
    .setFieldName(`Client Company`)
    .setTitle(`Client Company's Name`);

  var clientAddress = CardService.newTextInput()
    .setFieldName(`Client Address`)
    .setTitle(`Client Company's Address`);

  var dueDate = CardService.newDatePicker()
    .setFieldName('Due Date')
    .setTitle('Due Date')

  var paymentTerms = CardService.newTextInput()
    .setFieldName(`PayTerms`)
    .setTitle(`Warranty, returns policy...`);

  var totalTax = CardService.newTextInput()
    .setFieldName(`Tax`)
    .setTitle(`Total Tax (Optional)`);

  var discount = CardService.newTextInput()
    .setFieldName(`Discount`)
    .setTitle(`Discount (Optional)`);

  var email = CardService.newTextInput()
    .setFieldName(`Client Email`)
    .setTitle(`Client Email`);

  var postInvoice = CardService.newAction()
    .setFunctionName('viewInvoice');
  var newpostInvoiceButton = CardService.newTextButton()
    .setText('View Invoice')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(postInvoice);

  // send invoice to user.
  var invoiceName = CardService.newTextInput()
    .setFieldName('Invoice Name')
    .setTitle('Invoice Name');
  var sendInvoice = CardService.newAction()
    .setFunctionName('sndInvoice');
  var newInvoiceButton = CardService.newTextButton()
    .setText('Send Invoice')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(sendInvoice);

  invoiceSection.addWidget(contactName);
  invoiceSection.addWidget(clientName);
  invoiceSection.addWidget(email);
  invoiceSection.addWidget(clientAddress);
  invoiceSection.addWidget(dueDate);
  invoiceSection.addWidget(discount);
  invoiceSection.addWidget(totalTax);
  invoiceSection.addWidget(paymentTerms);
  invoiceSection.addWidget(CardService.newButtonSet().addButton(newpostInvoiceButton));

  sendInvoiceSection.addWidget(invoiceName);
  sendInvoiceSection.addWidget(CardService.newButtonSet().addButton(newInvoiceButton));

  var invcard = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("Create, Send and Track Invoices"))
    .addSection(invoiceSection)
    .addSection(sendInvoiceSection);

  let card = invcard.build();
  return card;
}


function sndInvoice(e) {
  var res = e['formInput'];
  var invoiceName = res['Invoice Name'] ? res['Invoice Name'] : 'Invoice';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = SpreadsheetApp.getActiveSpreadsheet().getId();

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() !== 'Invoicegen') {
      sheets[i].hideSheet()
    }
  }


  //or send as email
  var email = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B13").getValue();
  var company = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B2").getValue();
  var subject = `Invoice From ${company}`;
  var body = 'Invoice Ready';

  MailApp.sendEmail(email, subject, body, {
    attachments: [{
      fileName: invoiceName + ".pdf",
      content: ss.getBlob().getBytes(),
      mimeType: "application/pdf"
    }]
  })

  for (var i = 0; i < sheets.length; i++) {
    sheets[i].showSheet()
  }

  // send info to invoice template record.
  var date = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F10").getDisplayValue();
  var invoiceNum = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F9").getValue();
  var description = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B16:C16").getValue();
  var dueDate = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F11").getDisplayValue();
  var amount = SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F31").getValue();
  // var receiverAddr = email

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        date,
        invoiceNum,
        description,
        amount,
        dueDate,
        email,
        'no'
      ]
    ]
  }

  var optionalArgs = { valueInputOption: "USER_ENTERED" };
  Sheets.Spreadsheets.Values.append(
    request,
    ssID,
    'invoice!A:E',
    optionalArgs
  )


  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly sent invoice to ${email}`))
    .build() && invoice_card();
}


function onDrive() {
  return createFile();
}


function createFile() {
  var sheetName = CardService.newTextInput()
    .setFieldName('Sheet Name')
    .setTitle('Sheet Name');
  var createNewSheet = CardService.newAction()
    .setFunctionName('copyFile');
  var newSheetButton = CardService.newTextButton()
    .setText('Create New Sheet')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(createNewSheet);

  newSheetSection.addWidget(sheetName);
  newSheetSection.addWidget(CardService.newButtonSet().addButton(newSheetButton));

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Manage all bookkeeping in one place. Start by creating a Spreadsheet"))
    .addSection(newSheetSection)
    .build();
  return card;
}

function transaction() {
  // create three sections with 3 cards.
  // Sell
  var buttonAction = CardService.newAction()
    .setFunctionName('sellCard');
  transactionSection.addWidget(CardService.newDecoratedText()
    .setBottomLabel("Record Sales")
    .setEndIcon(CardService.newIconImage().setIconUrl('https://i.ibb.co/Ldk6ftd/Group-1-35.png'))
    .setText('Sell')
    .setOnClickAction(buttonAction));
  // Buy
  var buttonAction = CardService.newAction()
    .setFunctionName('buyCard');
  transactionSection.addWidget(CardService.newDecoratedText()
    .setBottomLabel("Restock Inventory")
    .setEndIcon(CardService.newIconImage().setIconUrl('https://i.ibb.co/NYFqrzK/Group-1-36.png'))
    .setText('Buy')
    .setOnClickAction(buttonAction));
  // Expense card
  var buttonAction = CardService.newAction()
    .setFunctionName('expenseCard');
  transactionSection.addWidget(CardService.newDecoratedText()
    .setBottomLabel("Record Outgoing Expenses")
    .setEndIcon(CardService.newIconImage().setIconUrl('https://i.ibb.co/NYFqrzK/Group-1-36.png'))
    .setText('Expense')
    .setOnClickAction(buttonAction));
  // Loan
  var buttonAction = CardService.newAction()
    .setFunctionName('loanCard');
  transactionSection.addWidget(CardService.newDecoratedText()
    .setBottomLabel("Record and Manage Liabilities Into Business")
    .setEndIcon(CardService.newIconImage().setIconUrl('https://www.linkpicture.com/q/icons8-forward-button-64.png'))
    .setText('Liability')
    .setOnClickAction(buttonAction));
  // Money In
  var buttonAction = CardService.newAction()
    .setFunctionName('innerflowCard');
  transactionSection.addWidget(CardService.newDecoratedText()
    .setBottomLabel("Record Incoming Funds")
    .setEndIcon(CardService.newIconImage().setIconUrl('https://i.ibb.co/Ldk6ftd/Group-1-35.png'))
    .setText('Money In')
    .setOnClickAction(buttonAction));
  // Money out
  var buttonAction = CardService.newAction()
    .setFunctionName('outCard');
  transactionSection.addWidget(CardService.newDecoratedText()
    .setBottomLabel("Pay for Liabilities")
    .setEndIcon(CardService.newIconImage().setIconUrl('https://i.ibb.co/NYFqrzK/Group-1-36.png'))
    .setText('Money Out')
    .setOnClickAction(buttonAction));

  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("Record all bookkeeping actions in your sheet").setImageUrl('https://www.linkpicture.com/q/32x32-google.png'))
    .addSection(transactionSection)
    .build();
  return card;
}

/*
// This is a statue for reference.
function transaction_card() {
  var description = CardService.newTextInput()
    .setFieldName('Description')
    .setTitle('Description');

  var amount = CardService.newTextInput()
    .setFieldName('Amount')
    .setTitle('Amount');

  var debit = CardService.newSelectionInput().setTitle('From')
    .setFieldName('Debit')
    .setType(CardService.SelectionInputType.DROPDOWN);

  INPUT_MAP.forEach((language, index, array) => {
    debit.addItem(language.text, language.val, language.val == true);
  })

  var credit = CardService.newSelectionInput().setTitle('To')
    .setFieldName('Credit')
    .setType(CardService.SelectionInputType.DROPDOWN);

  INPUT_MAP.forEach((language, index, array) => {
    credit.addItem(language.text, language.val, language.val == true);
  })

  inputSheetSection.addWidget(description);
  inputSheetSection.addWidget(amount);
  inputSheetSection.addWidget(debit);
  inputSheetSection.addWidget(credit);


  buttonSheetSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Record Transaction')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('submitRecord'))
      .setDisabled(false)));

  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("Record Transactions"))
    .addSection(inputSheetSection)
    .addSection(buttonSheetSection)
    .build();
  return card;
}
 */
function itemCard(price, payment_method, funcAction) {
  var item = CardService.newTextInput()
    .setFieldName('Item Name')
    .setTitle('Item Name');

  var description = CardService.newTextInput()
    .setFieldName('Description')
    .setTitle('Description');

  var amount = CardService.newTextInput()
    .setFieldName('Amount')
    .setTitle(price);

  var quantity = CardService.newTextInput()
    .setFieldName('Quantity')
    .setTitle('Quantity');

  var debit = CardService.newSelectionInput().setTitle(payment_method)
    .setFieldName('Debit')
    .setType(CardService.SelectionInputType.DROPDOWN);

  PAYMENT_MAP.forEach((language, index, array) => {
    debit.addItem(language.text, language.val, language.val == true);
  })

  inputSheetSection.addWidget(item);
  inputSheetSection.addWidget(description);
  inputSheetSection.addWidget(amount);
  inputSheetSection.addWidget(quantity);
  inputSheetSection.addWidget(debit);

  if (funcAction === 'sell') {
    buttonSheetSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Record Transaction')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('sellAction'))
        .setDisabled(false)));
  } else if (funcAction === 'buy') {
    buttonSheetSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Record Transaction')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('buyAction'))
        .setDisabled(false)));
  }

  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("Record Transactions"))
    .addSection(inputSheetSection)
    .addSection(buttonSheetSection)
    .build();
  return card;
}

function sellCard() {
  return itemCard(SELLING_PRICE, PAYMENT_METHOD, 'sell')
}

function sellAction(e) {
  var res = e['formInput'];

  var ItemName = res['Item Name'] ? res['Item Name'] : '';
  var Description = res['Description'] ? res['Description'] : '';
  var Amount = res['Amount'] ? res['Amount'] : '';
  var Quantity = res['Quantity'] ? res['Quantity'] : '';
  var Debit = res['Debit'] ? res['Debit'] : '';

  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;
  const transactiomNumber = Math.floor(100000 + Math.random() * 900000);

  // Add today's date
  // Add unique reference number

  let Total = Amount * Quantity;

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        ItemName,
        Description,
        Amount,
        Quantity,
        Total,
        Debit,
      ]
    ]
  }

  var optionalArgs = { valueInputOption: "USER_ENTERED" };
  Sheets.Spreadsheets.Values.append(
    request,
    spreadsheetId,
    'Sales!A:G',
    optionalArgs
  )

  var transrequest = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
        Description,
        Total,
        'Sales Account',
        Debit
      ]
    ]
  }

  Sheets.Spreadsheets.Values.append(
    transrequest,
    spreadsheetId,
    'Transactions!A:E',
    optionalArgs
  )

  // Double entry bookkeeping one.
  var bookRequestOne = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
        Description,
        ItemName,
        'Sales Account',
        Total
      ]
    ]
  }

  Sheets.Spreadsheets.Values.append(
    bookRequestOne,
    spreadsheetId,
    'General Ledger!A:I',
    optionalArgs
  )

  // Double entry bookkeeping two.
  var bookRequestTwo = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
        Description,
        ItemName,
        'Inventory Account',
        '',
        Total
      ]
    ]
  }

  Sheets.Spreadsheets.Values.append(
    bookRequestTwo,
    spreadsheetId,
    'General Ledger!A:I',
    optionalArgs
  )

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly Recorded Transaction`))
    .build() && sellCard();
}

function buyCard() {
  return itemCard(BUYING_PRICE, PAYMENT_METHOD, 'buy')
}

function buyAction(e) {
  var res = e['formInput'];

  var ItemName = res['Item Name'] ? res['Item Name'] : '';
  var Description = res['Description'] ? res['Description'] : '';
  var Amount = res['Amount'] ? res['Amount'] : '';
  var Quantity = res['Quantity'] ? res['Quantity'] : '';
  var Debit = res['Debit'] ? res['Debit'] : '';

  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;
  const transactiomNumber = Math.floor(100000 + Math.random() * 900000);

  // Add today's date
  // Add unique reference number

  let Total = Amount * Quantity;

  var optionalArgs = { valueInputOption: "USER_ENTERED" };

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
        Description,
        Total,
        Debit,
        'Inventory Account',
      ]
    ]
  }

  Sheets.Spreadsheets.Values.append(
    request,
    spreadsheetId,
    'Transactions!A:E',
    optionalArgs
  )

  var nav = CardService.newNavigation().pushCard(buyCard());

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly Recorded Sales`))
    .setNavigation(nav)
    .build();
}

function inflowCard(inflowTitle, inflowAction) {
  var inflowCategory = CardService.newSelectionInput().setTitle(inflowTitle)
    .setFieldName('Category')
    .setType(CardService.SelectionInputType.DROPDOWN);

  if (inflowAction === 'loan') {
    LIABILITY_MAP.forEach((language, index, array) => {
      inflowCategory.addItem(language.text, language.val, language.val == true);
    })
  } else if (inflowAction === 'in') {
    INPUT_MAP.forEach((language, index, array) => {
      inflowCategory.addItem(language.text, language.val, language.val == true);
    })
  } else if (inflowAction === 'out') {
    OUT_MAP.forEach((language, index, array) => {
      inflowCategory.addItem(language.text, language.val, language.val == true);
    })
  }

  var title = CardService.newTextInput()
    .setFieldName(`Vendor Name`)
    .setTitle(`Entity's Name`);

  var description = CardService.newTextInput()
    .setFieldName('Description')
    .setTitle('Description');

  var amount = CardService.newTextInput()
    .setFieldName('Amount')
    .setTitle('Amount');

  inputSheetSection.addWidget(inflowCategory);
  inputSheetSection.addWidget(title);
  inputSheetSection.addWidget(description);
  inputSheetSection.addWidget(amount);

  if (inflowAction === 'loan') {
    buttonSheetSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Record Transaction')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('loanAction'))
        .setDisabled(false)));
  } else if (inflowAction === 'in') {
    buttonSheetSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Record Transaction')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('inAction'))
        .setDisabled(false)));
  } else if (inflowAction === 'out') {
    buttonSheetSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Record Transaction')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('expAction'))
        .setDisabled(false)));
  }

  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("Record Transactions"))
    .addSection(inputSheetSection)
    .addSection(buttonSheetSection)
    .build();
  return card;
}

function expenseCard() {
  return inflowCard(EXP_CATEGORY, "out")
}

function expAction(e) {
  var res = e['formInput'];

  var ItemName = res['Item Name'] ? res['Item Name'] : '';
  var Description = res['Description'] ? res['Description'] : '';
  var Amount = res['Amount'] ? res['Amount'] : '';
  var Debit = res['Category'] ? res['Category'] : '';

  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;
  const transactiomNumber = Math.floor(100000 + Math.random() * 900000);

  // Add today's date
  // Add unique reference number

  var optionalArgs = { valueInputOption: "USER_ENTERED" };

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
        Description,
        Amount,
        'Bank',
        Debit
      ]
    ]
  }

  Sheets.Spreadsheets.Values.append(
    request,
    spreadsheetId,
    'Transactions!A:E',
    optionalArgs
  )

  var nav = CardService.newNavigation().pushCard(buyCard());

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly Recorded Sales`))
    .setNavigation(nav)
    .build();
}

function innerflowCard() {
  // An action response that opens a link in full screen
  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("This section is still a work in progress"))
    .build();
  return card;
}

function inAction() { }

function loanCard() {
  return inflowCard(LOAN_CATEGORY, "loan")
}

function loanAction(e) {
  var res = e['formInput'];

  var Description = res['Description'] ? res['Description'] : '';
  var Amount = res['Amount'] ? res['Amount'] : '';
  var Credit = res['Category'] ? res['Category'] : '';

  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;
  const transactiomNumber = Math.floor(100000 + Math.random() * 900000);

  // Add today's date
  // Add unique reference number

  var optionalArgs = { valueInputOption: "USER_ENTERED" };

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
        Description,
        Amount,
        Credit,
        'Bank',
      ]
    ]
  }

  Sheets.Spreadsheets.Values.append(
    request,
    spreadsheetId,
    'Transactions!A:E',
    optionalArgs
  )

  var nav = CardService.newNavigation().pushCard(loanCard());

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly Recorded Sales`))
    .setNavigation(nav)
    .build();
}

function outCard() {
  var card = CardService.newCardBuilder()
    .setName("Card name")
    .setHeader(CardService.newCardHeader().setTitle("This section is still a work in progress"))
    .build();
  return card;
}

function template() {
  var currentButton = CardService.newAction()
    .setFunctionName('invoice_card');
  var newSheetButton = CardService.newTextButton()
    .setText('Use Default Template')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(currentButton);

  selectInvoiceSection.addWidget(CardService.newButtonSet().addButton(newSheetButton));

  var selectedGrid = CardService.newGrid()
    .setTitle("Selected Template")
    .setNumColumns(2)
    .addItem(CardService.newGridItem()
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img1-peppu.jpg')));
  selectInvoiceSection.addWidget(selectedGrid);

  /*if () {
    selectedGrid = CardService.newGrid()
.setTitle("My Grid")
.setNumColumns(2)
.addItem(CardService.newGridItem()
    .setTitle("My item")
    .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/Transaction.png')));
    selectInvoiceSection.addWidget(selectedGrid);
  }
  */

  var grid = CardService.newGrid()
    .setTitle("Choose Template")
    .setBorderStyle(CardService.newBorderStyle().setType(CardService.BorderType.STROKE))
    .setOnClickAction(
      CardService.newAction()
        .setFunctionName("testGrid"))
    .setNumColumns(2)
    .addItem(CardService.newGridItem()
      .setTitle("X1")
      .setIdentifier("001")
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img1-peppu.jpg')))
    .addItem(CardService.newGridItem()
      .setTitle("X2")
      .setIdentifier("002")
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img2-peppu.jpg')))
    .addItem(CardService.newGridItem()
      .setTitle("Y3")
      .setIdentifier("003")
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/Invoice-Template-3.jpg')))
    .addItem(CardService.newGridItem()
      .setTitle("Y4")
      .setIdentifier("004")
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img4-peppu.jpg')))

  selectInvoiceSection.addWidget(grid);

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Manage all bookkeeping in one place. Start by creating a Spreadsheet"))
    .addSection(selectInvoiceSection)
    .build();
  return card;

}


function onSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getName() == 'Instructions') {
    var buttonAction = CardService.newAction()
      .setFunctionName('template');
    navigationSection.addWidget(CardService.newDecoratedText()
      .setBottomLabel("Create, Send and Track Invoices")
      .setEndIcon(CardService.newIconImage().setIconUrl('https://www.linkpicture.com/q/icons8-forward-button-64.png'))
      .setText('Invoice Actions')
      .setOnClickAction(buttonAction));

    var buttonAction = CardService.newAction()
      .setFunctionName('transaction');
    navigationSection.addWidget(CardService.newDecoratedText()
      .setBottomLabel("Record Transactions in Sheet")
      .setEndIcon(CardService.newIconImage().setIconUrl('https://www.linkpicture.com/q/icons8-forward-button-64.png'))
      .setText('Transaction Actions')
      .setOnClickAction(buttonAction));

    var card = CardService.newCardBuilder()
      .setName("Card name")
      .setHeader(CardService.newCardHeader().setTitle("Perform all bookkeeping actions in your sheet").setImageUrl('https://www.linkpicture.com/q/32x32-google.png'))
      .addSection(navigationSection)
      .build();
    return card;
  } else {
    newSheetSection.addWidget(CardService.newDecoratedText().setText(`Heyya!! You can't use PayTrack with this Spreadsheet. Please create a new Spreadsheet and start accounting from the new Sheet. Contact peppubooks@gmail.com if you have any challenge.`).setWrapText(true))

    var sheetName = CardService.newTextInput()
      .setFieldName('Sheet Name')
      .setTitle('Sheet Name');
    var createNewSheet = CardService.newAction()
      .setFunctionName('copySheet');
    var newSheetButton = CardService.newTextButton()
      .setText('Create New Sheet')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(createNewSheet);

    newSheetSection.addWidget(sheetName);
    newSheetSection.addWidget(CardService.newButtonSet().addButton(newSheetButton));

    var card = CardService.newCardBuilder()
      .addSection(newSheetSection)
      .build();
    return card;
  }
}

function copyFile(e) {
  var res = e['formInput'];
  var sheetName = res['Sheet Name'] ? res['Sheet Name'] : '';
  let id = '1S4GMiZ0H0_6OHH7DEnjZt07-6kk0eMP4YSNUmRcKZXA';
  let file = Drive.Files.copy({ title: sheetName }, id);
  let new_file_id = file.id;
  // An action response that opens the spreadsheet in full screen
  var nav = CardService.newNavigation().pushCard(createFile());
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly Created File ${sheetName}. Open file from your Drive or allow redirect from drive.`)).setNavigation(nav)
    .setOpenLink(CardService.newOpenLink()
      .setUrl(`https://docs.google.com/spreadsheets/d/${new_file_id}`)
      .setOpenAs(CardService.OpenAs.FULL_SIZE))
    .build();


}

function copySheet(e) {
  var res = e['formInput'];
  var sheetName = res['Sheet Name'] ? res['Sheet Name'] : '';
  let id = '1S4GMiZ0H0_6OHH7DEnjZt07-6kk0eMP4YSNUmRcKZXA';
  let file = Drive.Files.copy({ title: sheetName }, id);
  let new_file_id = file.id;
  // An action response that opens the spreadsheet in full screen
  return openUrl(`https://docs.google.com/spreadsheets/d/${new_file_id}`);
}

function openUrl(url) {
  var html = HtmlService.createHtmlOutput('<!DOCTYPE html><html>'
    // Offer URL as clickable link in case above code fails.
    + '<body style="word-break:break-word;font-family:sans-serif;">Failed to open spreadsheet automatically.  <br/><a href="' + url + '" target="_blank" onclick="window.close()">Click here to open spreadsheet</a>.</body>'
    + '<script>google.script.host.setHeight(55);google.script.host.setWidth(410)</script>'
    + '</html>')
    .setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening ...");
}

function viewInvoice(e) {
  var res = e['formInput'];
  var contactName = res['Contact Name'] ? res['Contact Name'] : '';
  var clientName = res['Client Company'] ? res['Client Company'] : '';
  var clientEmail = res['Client Email'] ? res['Client Email'] : '';
  var clientAddress = res['Client Address'] ? res['Client Address'] : '';
  var dueDate = res['Due Date'] ? res['Due Date'] : '';
  var payTerms = res['PayTerms'] ? res['PayTerms'] : '';
  var discount = res['Discount'] ? res['Discount'] : 0;
  var totalTax = res['totalTax'] ? res['totalTax'] : 0;
  const invNumber = Math.floor(100000 + Math.random() * 900000);

  let date = dueDate.msSinceEpoch;
  // WE NEED TO RETRIEVE USER'S TIMEZONE
  let formatDate = Utilities.formatDate(new Date(date), "GMT", "yyyy/MM/dd")

  // set today's date as invoice.
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;

  var coyName = SpreadsheetApp.getActiveSheet().getRange("Instructions!C11:I11").getValue();
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B2:D2").setValue(coyName);
  var coyAddr = SpreadsheetApp.getActiveSheet().getRange("Instructions!C12:I12").getValue();
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B3:D3").setValue(coyAddr);
  var phoneNum = SpreadsheetApp.getActiveSheet().getRange("Instructions!C13:I13").getValue();
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B5:D5").setValue(phoneNum);
  var bankName = SpreadsheetApp.getActiveSheet().getRange("Instructions!C14:I14").getValue();
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!C35:F35").setValue(bankName);
  var acctNum = SpreadsheetApp.getActiveSheet().getRange("Instructions!C15:I15").getValue();
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!C37:F37").setValue(acctNum);
  var acctName = SpreadsheetApp.getActiveSheet().getRange("Instructions!C16:I16").getValue();
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!C36:F36").setValue(acctName);
  var routingNum = SpreadsheetApp.getActiveSheet().getRange("Instructions!C17:I17").getValue();
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!C38:F38").setValue(routingNum);
  var paymentMthd = SpreadsheetApp.getActiveSheet().getRange("Instructions!C18:I18").getValue();
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!C39:F39").setValue(paymentMthd)
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B10").setValue(contactName);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B11").setValue(clientName);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B12").setValue(clientAddress);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B13").setValue(clientEmail);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F11").setValue(formatDate);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F26").setValue(discount);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F28").setValue(totalTax);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F9").setValue(`INV${invNumber}`);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!B35:F35").setValue(payTerms);
  SpreadsheetApp.getActiveSheet().getRange("Invoicegen!F10").setValue(formattedToday);

  return invoice_card();
}

function submitRecord(e) {
  var res = e['formInput'];


  var Description = res['Description'] ? res['Description'] : '';
  var Amount = res['Amount'] ? res['Amount'] : '';
  var Debit = res['Debit'] ? res['Debit'] : '';
  var Credit = res['Credit'] ? res['Credit'] : '';


  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;
  const transactiomNumber = Math.floor(100000 + Math.random() * 900000);

  // Add today's date
  // Add unique reference number

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
        Description,
        Amount,
        Debit,
        Credit
      ]
    ]
  }

  var optionalArgs = { valueInputOption: "USER_ENTERED" };
  Sheets.Spreadsheets.Values.append(
    request,
    spreadsheetId,
    'Transactions!A:E',
    optionalArgs
  )

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Successfuly Recorded Transaction`))
    .build() && transaction_card();

}

function testGrid(e) {
  let val = e.parameters.grid_item_identifier;
  if (val == '001') {
    var selectedGrid = CardService.newGrid()
      .setTitle("Selected Template")
      .setNumColumns(2)
      .addItem(CardService.newGridItem()
        .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img1-peppu.jpg')));
    selectInvoiceSection.addWidget(selectedGrid);

    var buttonAction = CardService.newAction()
      .setFunctionName('copyTemplateOne');
    selectInvoiceSection.addWidget(CardService.newTextButton()
      .setText('Create Template')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(buttonAction))
  } else if (val == '002') {
    var selectedGrid = CardService.newGrid()
      .setTitle("Selected Template")
      .setNumColumns(2)
      .addItem(CardService.newGridItem()
        .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img2-peppu.jpg')));
    selectInvoiceSection.addWidget(selectedGrid);

    var buttonAction = CardService.newAction()
      .setFunctionName('copyTemplateTwo');
    selectInvoiceSection.addWidget(CardService.newTextButton()
      .setText('Create Template')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(buttonAction))
  }
  else if (val == '003') {
    var selectedGrid = CardService.newGrid()
      .setTitle("Selected Template")
      .setNumColumns(2)
      .addItem(CardService.newGridItem()
        .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/Invoice-Template-3.jpg')));
    selectInvoiceSection.addWidget(selectedGrid);

    var buttonAction = CardService.newAction()
      .setFunctionName('copyTemplateThree');
    selectInvoiceSection.addWidget(CardService.newTextButton()
      .setText('Create Template')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(buttonAction))
  } else if (val == '004') {
    var selectedGrid = CardService.newGrid()
      .setTitle("Selected Template")
      .setNumColumns(2)
      .addItem(CardService.newGridItem()
        .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img4-peppu.jpg')));
    selectInvoiceSection.addWidget(selectedGrid);

    var buttonAction = CardService.newAction()
      .setFunctionName('copyTemplateFour');
    selectInvoiceSection.addWidget(CardService.newTextButton()
      .setText('Create Template')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(buttonAction))
  }

  var grid = CardService.newGrid()
    .setTitle("Choose Template")
    .setBorderStyle(CardService.newBorderStyle().setType(CardService.BorderType.STROKE))
    .setOnClickAction(
      CardService.newAction()
        .setFunctionName("testGrid"))
    .setNumColumns(2)
    .addItem(CardService.newGridItem()
      .setTitle("X1")
      .setIdentifier("001")
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img1-peppu.jpg')))
    .addItem(CardService.newGridItem()
      .setTitle("X2")
      .setIdentifier("002")
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img2-peppu.jpg')))
    .addItem(CardService.newGridItem()
      .setTitle("Y3")
      .setIdentifier("003")
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/Invoice-Template-3.jpg')))
    .addItem(CardService.newGridItem()
      .setTitle("Y4")
      .setIdentifier("004")
      .setImage(CardService.newImageComponent().setImageUrl('https://www.linkpicture.com/q/img4-peppu.jpg')))

  selectInvoiceSection.addWidget(grid);

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Manage all bookkeeping in one place. Start by creating a Spreadsheet"))
    .addSection(selectInvoiceSection)
    .build();
  return card;
}

var cardBuilder1 = CardService.newCardBuilder();
var cardBuilder2 = CardService.newCardBuilder();

// copyTemplateOne to copy first invoice template
function copyTemplateOne() {

  var source = SpreadsheetApp.getActiveSpreadsheet();

  var destination = SpreadsheetApp.openById("1yRfnRiEGX9LmkkaqFFEnouQ1LyBJJkOGxgeYUsepEHg");
  let sheet = destination.getSheets()[0];

  // delete invoicegen file
  let invSheet = source.getSheetByName('Invoicegen');
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(invSheet);

  // copy to PayTrack template
  sheet.copyTo(source).setName('Invoicegen');
  // call invoice card.
  return invoice_card();
}

function copyTemplateTwo() {
  var source = SpreadsheetApp.getActiveSpreadsheet();

  var destination = SpreadsheetApp.openById("1yRfnRiEGX9LmkkaqFFEnouQ1LyBJJkOGxgeYUsepEHg");
  let sheet = destination.getSheets()[1];

  // delete invoicegen file
  let invSheet = source.getSheetByName('Invoicegen');
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(invSheet);

  // copy to PayTrack template
  sheet.copyTo(source).setName('Invoicegen');

  // call invoice card.
  return invoice_card();
}

function copyTemplateThree() {
  var source = SpreadsheetApp.getActiveSpreadsheet();

  var destination = SpreadsheetApp.openById("1yRfnRiEGX9LmkkaqFFEnouQ1LyBJJkOGxgeYUsepEHg");
  let sheet = destination.getSheets()[2];

  // delete invoicegen file
  let invSheet = source.getSheetByName('Invoicegen');
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(invSheet);

  // copy to PayTrack template
  sheet.copyTo(source).setName('Invoicegen');

  // call invoice card.
  return invoice_card();
}

function copyTemplateFour() {
  var source = SpreadsheetApp.getActiveSpreadsheet();

  var destination = SpreadsheetApp.openById("1yRfnRiEGX9LmkkaqFFEnouQ1LyBJJkOGxgeYUsepEHg");
  let sheet = destination.getSheets()[3];

  // delete invoicegen file
  let invSheet = source.getSheetByName('Invoicegen');
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(invSheet);

  // copy to PayTrack template
  sheet.copyTo(source).setName('Invoicegen');

  // call invoice card.
  return invoice_card();
}

// function to send email everyday about due invoice to user's client.
function sendMail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoice')

  // get range of data and values, containing email, subject, etc.
  var dataRange = sheet.getRange(3, 5, 990, 3);
  var data = dataRange.getValues();

  // loop through to get all values
  for (i in data) {
    var row = data[i];
    let tDate = new Date();
    let cDate = new Date(row[0]);
    let diff = Math.abs(tDate.valueOf() - cDate.valueOf());
    let difference = (diff / (1000 * 60 * 60 * 24));
    if (row[2] == 'no' && difference <= 3) {
      var emailAddress = row[1];
      var subject = 'Invoice Due';
      var date = row[0];
      /* Change PayTrack to the name of the user or their firm. */
      var coyName = SpreadsheetApp.getActiveSheet().getRange("Instructions!C11:I11").getValue();
      // You have an unpaid invoice for {} via paytrack.
      // you have few days left to pay the invoice as it is due on ...
      // Your invoice is due today
      // Your invoice is past the due date.
      var message = `You have an unpaid invoice for ${coyName} via PayTrack (https://workspace.google.com/marketplace/app/paytrack/913987535189) due for ${date}`;
      try {
        // send email to client.
        MailApp.sendEmail(emailAddress, subject, message);
      } catch (errorDetails) {
        Logger.log(errorDetails);
      }
    }
  }
}

// completeTransaction function contains all triggers for payTrack
function completeTransaction() {
  // Trigger to send mails to user's client.
  ScriptApp.newTrigger('sendMail')
    .timeBased()
    .atHour(7)
    .nearMinute(0)
    .everyDays(1)
    .create();

  // Trigger to complete transaction after invoice is marked as paid
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('completetrans')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Since this trigger should be installed once, this notification
  // lets the user know that they've triggered it successfully.
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Hi, you've successfully installed triggers, do not click on this button again`))
    .build();
}

// `completetrans` function to complete transaction after user marks invoice as paid
function completetrans() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Format date for transaction date
  const today = new Date();
  const yyyy = today.getFullYear();
  let mm = today.getMonth() + 1; // Months start at 0!
  let dd = today.getDate();

  if (dd < 10) dd = '0' + dd;
  if (mm < 10) mm = '0' + mm;

  const formattedToday = yyyy + '/' + mm + '/' + dd;

  // Create random number for transaction number
  const transactiomNumber = Math.floor(100000 + Math.random() * 900000);

  var row = ss.getActiveRange().getRow();
  // Data for append request
  var Description = SpreadsheetApp.getActiveSheet().getRange(`Invoice!C${row}`).getValue();
  var Amount = SpreadsheetApp.getActiveSheet().getRange(`Invoice!D${row}`).getValue();
  var Debit = 'Cash';
  var Credit = 'Bank';

  var optionalArgs = { valueInputOption: "USER_ENTERED" };

  var request = {
    "majorDimension": "ROWS",
    "values": [
      [
        formattedToday,
        `TRAN${transactiomNumber}`,
        Description,
        Amount,
        Debit,
        Credit
      ]
    ]
  }

  // if statement, to append value if transaction has been completed and SheetName of edit is invoice sheet.
  if (ss.getActiveSheet().getSheetName() == 'Invoice' && ss.getActiveRange().getValue() == 'yes') {
    try {
      Sheets.Spreadsheets.Values.append(
        request,
        ss.getId(),
        'Transactions!A:E',
        optionalArgs
      )
    } catch (errorDetails) {
      CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText(errorDetails))
        .build();
    }
  }
}

function onDocs() {
  // An action response that opens a documentation in full screen
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl("https://docs.peppubooks.com")
      .setOpenAs(CardService.OpenAs.FULL_SIZE))
    .build();
}