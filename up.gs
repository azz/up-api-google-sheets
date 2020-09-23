/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

const TOKEN_CACHE_DURATION_SECONDS = 60 * 60;
const TOKEN_CACHE_DURATION_HUMAN = "1 hour";

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getActive().addMenu("Up API", [
    { name: "Set Up...", functionName: "init_" },
    null,
    { name: "UP_PING", functionName: "insertUpPing_" },
    { name: "UP_ACCOUNTS", functionName: "insertUpAccounts_" },
    { name: "UP_TRANSACTIONS", functionName: "insertUpTransactions_" },
  ]);
}

function init_() {
  const sheet = SpreadsheetApp.getActiveSheet().setName("Up API");
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    "Up API",
    "Enter your Up API Personal Access Token. This token will be cached for " +
      TOKEN_CACHE_DURATION_HUMAN +
      ".",
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  CacheService.getUserCache().put(
    "UP_API_TOKEN",
    result.getResponseText(),
    TOKEN_CACHE_DURATION_SECONDS
  );

  insert_("=UP_PING()", 1, sheet.getRange("A1"));
}

function insert_(formula, numberOfColumns, range) {
  const sheet = SpreadsheetApp.getActiveSheet();
  range = range || sheet.getActiveRange();
  range.offset(0, 0, 1, 1).setValue(formula);

  const headingRange = range.offset(0, 0, 1, numberOfColumns);
  headingRange
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setForegroundColor("#1a1a22")
        .setBold(true)
        .build()
    )
    .setBackground("#ff7a64")
    .setHorizontalAlignment("center")
    .activate();

  SpreadsheetApp.flush();
  sheet.autoResizeColumns(
    range.getColumn(),
    range.getColumn() + numberOfColumns
  );
}

function insertUpPing_() {
  insert_("=UP_PING()", 1);
}
function insertUpTransactions_() {
  insert_("=UP_TRANSACTIONS(false, 100)", UP_TRANSACTIONS_HEADINGS.length);
}
function insertUpAccounts_() {
  insert_("=UP_ACCOUNTS(100)", UP_ACCOUNTS_HEADINGS.length);
}

const UP_TRANSACTIONS_HEADINGS = [
  "Created At",
  "Settled At",
  "Status",
  "Direction",
  "Currency",
  "Value",
  "Description",
  "Category",
  "Parent Category",
  "Tags",
  "Message",
];

/**
 * @param {boolean} onlyDebits Only return debits
 * @param {number} input Page Size
 * @return Up Transactions
 * @customfunction
 */
function UP_TRANSACTIONS(onlyDebits = false, pageSize = 100) {
  return up_(`transactions?page[size]=${pageSize}`, (response) => {
    const transactions = onlyDebits
      ? response.data.filter((tx) => tx.attributes.amount.valueInBaseUnits < 0)
      : response.data;

    const table = transactions.map((transaction) => {
      const attributes = transaction.attributes;
      return [
        new Date(attributes.createdAt),
        attributes.settledAt ? new Date(attributes.settledAt) : "",
        attributes.status,
        attributes.amount.valueInBaseUnits < 0 ? "DEBIT" : "CREDIT",
        attributes.amount.currencyCode,
        Math.abs(Number(attributes.amount.value)),
        attributes.description,
        transaction.relationships.category.data
          ? transaction.relationships.category.data.id
          : "",
        transaction.relationships.parentCategory.data
          ? transaction.relationships.parentCategory.data.id
          : "",
        transaction.relationships.tags.data.map((tag) => tag.id).join(","),
        attributes.message,
      ];
    });
    return [UP_TRANSACTIONS_HEADINGS, ...table];
  });
}

const UP_ACCOUNTS_HEADINGS = [
  "Created At",
  "Type",
  "Name",
  "Currency",
  "Balance",
];

/**
 * @param {number} input Page Size
 * @return Up Accounts
 * @customfunction
 */
function UP_ACCOUNTS(pageSize = 100) {
  return up_(`accounts?page[size]=${pageSize}`, (response) => {
    const table = response.data.map((account) => {
      const attributes = account.attributes;
      return [
        new Date(attributes.createdAt),
        attributes.accountType,
        attributes.displayName,
        attributes.balance.currencyCode,
        attributes.balance.value,
      ];
    });
    return [UP_ACCOUNTS_HEADINGS, ...table];
  });
}

/**
 * @return Up Ping
 * @customfunction
 */
function UP_PING() {
  return up_(
    `util/ping`,
    (response) => ["Up API Status", response.meta.statusEmoji],
    (error) => [
      "Up API Status",
      error.message.includes("401") ? "Invalid Token" : error.message,
    ]
  );
}

function up_(url, tabulate, handleError = (e) => e.message) {
  const token = CacheService.getUserCache().get("UP_API_TOKEN");
  if (!token) {
    throw new Error('Please navigate to "Up API" → "Set Up..."');
  }

  try {
    let json = UrlFetchApp.fetch(`https://api.up.com.au/api/v1/${url}`, {
      headers: { Authorization: `Bearer ${token}` },
    }).getContentText();
    return tabulate(JSON.parse(json));
  } catch (error) {
    return handleError(error);
  }
}
