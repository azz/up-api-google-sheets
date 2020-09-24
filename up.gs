/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

const TOKEN_CACHE_DURATION_SECONDS = 60 * 60;
const TOKEN_CACHE_DURATION_HUMAN = "1 hour";
const MAX_RECORDS = 1000;

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Up API")
    .addItem("Set Up...", "init_")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Transactions")
        .addItem("All Transactions", "insertUpTransactions_")
        .addItem("Transactions for Account", "insertTransactionsForAccount_")
    )
    .addSubMenu(
      ui.createMenu("Accounts").addItem("All Accounts", "insertUpAccounts_")
    )
    .addSubMenu(
      ui
        .createMenu("Categories")
        .addItem("All Categories", "insertUpCategories_")
    )
    .addSubMenu(ui.createMenu("Tags").addItem("All Tags", "insertUpTags_"))
    .addSubMenu(ui.createMenu("Utilities").addItem("Ping", "insertUpPing_"))
    .addToUi();
}

function init_() {
  const sheet = SpreadsheetApp.getActiveSheet();
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

  // Force a recalculation every hour (and re-authentication when appropriate)
  SpreadsheetApp.getActive().setRecalculationInterval(
    SpreadsheetApp.RecalculationInterval.HOUR
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
function insertUpTags_() {
  insert_("=UP_TAGS()", UP_TAGS_HEADINGS.length);
}
function insertUpTransactions_() {
  insert_("=UP_TRANSACTIONS()", UP_TRANSACTIONS_HEADINGS.length);
}
function insertUpTransactionsForAccount_() {
  insert_("=UP_TRANSACTIONS_FOR_ACCOUNT()", UP_TRANSACTIONS_HEADINGS.length);
}
function insertUpAccounts_() {
  insert_("=UP_ACCOUNTS()", UP_ACCOUNTS_HEADINGS.length);
}
function insertUpCategories_() {
  insert_("=UP_CATEGORIES()", UP_CATEGORIES_HEADINGS.length);
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
 * @param {string} filterQuery The filter querystring to use, e.g. "filter[status]=HELD&filter[category]=booze"
 * @param {string} type 'ALL', 'DEBIT', 'CREDIT'
 * @return Up Transactions
 * @customfunction
 */
function UP_TRANSACTIONS(filterQuery = "", type = "ALL") {
  return up_(`transactions?${hackyUriEncode_(filterQuery)}`, {
    tabulate: (data) => tabulateTransactions_(type, data),
  });
}

/**
 * @param {string} accountId Up account's ID
 * @param {string} filterQuery The filter querystring to use, e.g. "filter[status]=HELD&filter[category]=booze"
 * @param {string} type 'ALL', 'DEBIT', 'CREDIT'
 * @return Up Transactions for Account
 * @customfunction
 */
function UP_TRANSACTIONS_FOR_ACCOUNT(
  accountId,
  filterQuery = "",
  type = "ALL"
) {
  if (!accountId) return "accountId is required.";
  return up_(
    `accounts/${accountId}/transactions?${hackyUriEncode_(filterQuery)}`,
    {
      tabulate: (data) => tabulateTransactions_(type, data),
    }
  );
}

function tabulateTransactions_(type, transactions) {
  if (type === "DEBIT")
    transactions = transactions.filter(
      (tx) => tx.attributes.amount.valueInBaseUnits < 0
    );
  if (type === "CREDIT")
    transactions = transactions.filter(
      (tx) => tx.attributes.amount.valueInBaseUnits > 0
    );

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
}

const UP_ACCOUNTS_HEADINGS = [
  "Account ID",
  "Created At",
  "Type",
  "Name",
  "Currency",
  "Balance",
];

/**
 * @return Up Accounts
 * @customfunction
 */
function UP_ACCOUNTS() {
  return up_(`accounts`, {
    tabulate(data) {
      const table = data.map((account) => {
        const attributes = account.attributes;
        return [
          account.id,
          new Date(attributes.createdAt),
          attributes.accountType,
          attributes.displayName,
          attributes.balance.currencyCode,
          attributes.balance.value,
        ];
      });
      return [UP_ACCOUNTS_HEADINGS, ...table];
    },
  });
}

const UP_CATEGORIES_HEADINGS = [
  "Category ID",
  "Category Name",
  "Parent Category ID",
];

/**
 * @return Up Categories
 * @customfunction
 */
function UP_CATEGORIES() {
  return up_(`categories`, {
    tabulate(data) {
      const table = data.map((category) => [
        category.id,
        category.attributes.name,
        category.relationships.parent.data
          ? category.relationships.parent.data.id
          : "all",
      ]);
      return [UP_CATEGORIES_HEADINGS, ...table, ["all", "All", ""]];
    },
  });
}

const UP_TAGS_HEADINGS = ["Tag"];

/**
 * @return Up Tags
 * @customfunction
 */
function UP_TAGS() {
  return up_(`tags`, {
    tabulate(data) {
      const table = data.map((tag) => [tag.id]);
      return [UP_TAGS_HEADINGS, ...table];
    },
  });
}

/**
 * @return Up Ping
 * @customfunction
 */
function UP_PING() {
  return up_(`util/ping`, {
    paginate: false,
    tabulate: (response) => ["Up API Status", response.meta.statusEmoji],
  });
}

function up_(path, { paginate = true, tabulate }) {
  const token = CacheService.getUserCache().get("UP_API_TOKEN");
  if (!token) {
    throw new Error('Please navigate to "Up API" → "Set Up..."');
  }

  try {
    let url = `https://api.up.com.au/api/v1/${path}`;
    let data = [];
    do {
      const json = UrlFetchApp.fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
        muteHttpExceptions: true,
      }).getContentText();
      const response = JSON.parse(json);
      if (response.errors) {
        return [["API Error"]].concat(
          response.errors.map((error) => [
            error.status,
            error.title,
            error.detail,
          ])
        );
      }
      if (!paginate) {
        return tabulate(response);
      }
      url = response.links ? response.links.next : null;
      data.push(...response.data);
    } while (url && data.length < MAX_RECORDS);
    return tabulate(data);
  } catch (error) {
    return ["ERROR", error.message];
  }
}

/* 🙈 */
function hackyUriEncode_(query) {
  return query
    .split("&")
    .map((kv) => {
      const [k, v] = kv.split("=");
      return `${k}=${encodeURIComponent(v)}`;
    })
    .join("&");
}
