/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/* global CacheService, UrlFetchApp, SpreadsheetApp */
/* eslint-disable no-unused-vars */
/* eslint-disable max-params */

const {ThemeColorType, RecalculationInterval} = SpreadsheetApp;

const TOKEN_CACHE_DURATION_SECONDS = 60 * 60 * 24;
const TOKEN_CACHE_DURATION_HUMAN = '1 day';
const MAX_RECORDS = 200;

const APP_NAME = 'Up API';

const DARK_GREEN = '#385454';
const DARK_BLUE = '#242430';
const X_DARK_BLUE = '#1A1A22';
const API_BLUE = '#3EA9F5';
const ALT_WHITE = '#FBFBFA';
const BRAND_ORANGE = '#FF7A64';
const BRAND_YELLOW = '#FFF06B';
const BRAND_BLUE = '#4E6280';
const LOGO_BLUE = '#3EA9F5';
const BRAND_PINK = '#FF8BB5';
const BRAND_GREEN = '#305555';
const AQUA = '#25BBB8';
const AMOUNT_GREEN = '#00BC83';
const RED = '#EF3B3D';
const GREY = '#D2D2D2';
const ANOTHER_GREY = '#A4A4A8';
const DARK_GREY = '#34333B';
const YELLOW = '#FFEF6B';
const YELLOW_LIGHT = '#FFFCE2';
const WHITE = '#FFFFFF';

const THEME = new Map([
  [ThemeColorType.BACKGROUND, BRAND_YELLOW],
  [ThemeColorType.TEXT, DARK_BLUE],
  [ThemeColorType.ACCENT1, BRAND_ORANGE],
  [ThemeColorType.ACCENT2, BRAND_BLUE],
  [ThemeColorType.ACCENT3, BRAND_GREEN],
  [ThemeColorType.ACCENT4, YELLOW_LIGHT],
  [ThemeColorType.ACCENT5, BRAND_PINK],
  [ThemeColorType.ACCENT6, AQUA],
  [ThemeColorType.HYPERLINK, LOGO_BLUE],
]);

/*
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu(APP_NAME)
    .addItem('Set Up...', 'init_')
    .addItem('Log out', 'logOut_')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('Transactions')
        .addItem('All Transactions', 'insertUpTransactions_')
        .addItem('Transactions between dates', 'insertUpTransactionsBetween_')
        .addItem('Transactions for Account', 'insertUpTransactionsForAccount_'),
    )
    .addSubMenu(
      ui.createMenu('Accounts').addItem('All Accounts', 'insertUpAccounts_'),
    )
    .addSubMenu(
      ui
        .createMenu('Categories')
        .addItem('All Categories', 'insertUpCategories_'),
    )
    .addSubMenu(ui.createMenu('Tags').addItem('All Tags', 'insertUpTags_'))
    .addSubMenu(ui.createMenu('Utilities').addItem('Ping', 'insertUpPing_'))
    .addToUi();
}

function init_() {
  const doc = SpreadsheetApp.getActive();
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // TODO: after the Up API supports, OAuth2, we won't need to use tokens!
  // https://developers.google.com/gsuite/add-ons/how-tos/non-google-services

  const result = ui.prompt(
    APP_NAME,
    'Enter your Up API Personal Access Token.\n' +
      'You can retrieve this from https://api.up.com.au.\n\n' +
      'You will be logged in for ' +
      TOKEN_CACHE_DURATION_HUMAN +
      '. After this time, your data will be cleared and you must provide your token again.',
    ui.ButtonSet.OK_CANCEL,
  );
  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  TokenCache.store(result.getResponseText());

  const theme = SpreadsheetApp.getActive().getSpreadsheetTheme();
  for (const [key, value] of THEME.entries()) {
    theme.setConcreteColor(
      key,
      SpreadsheetApp.newColor().setRgbColor(value).build(),
    );
  }

  // Force a recalculation every hour (and re-authentication when appropriate)
  doc.setRecalculationInterval(RecalculationInterval.HOUR);

  const statusRange = sheet.getRange('A1:B2');
  if (statusRange.isBlank()) {
    insert_('=UP_PING()', 1, statusRange);
    // Define a named range we can use to force other formulas to recalculate
    doc.setNamedRange('Yeah', statusRange);
  }
}

class TokenCache {
  static get cache() {
    return CacheService.getUserCache();
  }

  static store(token, expiry = TOKEN_CACHE_DURATION_SECONDS) {
    const expiryDate = new Date();
    expiryDate.setSeconds(expiryDate.getSeconds() + expiry);
    this.cache.put('token', token, expiry);
    this.cache.put('tokenExpiry', expiryDate.toISOString(), expiry);
  }

  static retrieve() {
    return this.cache.getAll(['token', 'tokenExpiry']);
  }

  static expire() {
    const hadToken = Boolean(this.cache.get('token'));
    this.cache.removeAll(['token', 'tokenExpiry']);
    return hadToken;
  }
}

function logOut_() {
  const ui = SpreadsheetApp.getUi();
  if (TokenCache.expire()) {
    ui.alert(
      APP_NAME,
      'You have been successfully logged out.',
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert(APP_NAME, 'You are not currently logged in.', ui.ButtonSet.OK);
  }
}

function insert_(formula, numberOfColumns, range) {
  const sheet = SpreadsheetApp.getActiveSheet();
  range = range || sheet.getActiveRange();
  range.offset(0, 0, 1, 1).setValue(formula);

  const headingRange = range.offset(0, 0, 1, numberOfColumns);
  headingRange
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setForegroundColor(X_DARK_BLUE)
        .setBold(true)
        .build(),
    )
    .setBackground(BRAND_ORANGE)
    .activate();

  SpreadsheetApp.flush();
  sheet.autoResizeColumns(
    range.getColumn(),
    range.getColumn() + numberOfColumns,
  );
}

function insertUpPing_() {
  insert_('=UP_PING()', 2);
}

function insertUpTags_() {
  insert_('=UP_TAGS(Yeah)', UP_TAGS_HEADINGS.length);
}

function insertUpTransactions_() {
  insert_('=UP_TRANSACTIONS(Yeah)', UP_TRANSACTIONS_HEADINGS.length);
}

function insertUpTransactionsBetween_() {
  insert_(
    '=UP_TRANSACTIONS_BETWEEN(Yeah, TODAY() - 30, TODAY())',
    UP_TRANSACTIONS_HEADINGS.length,
  );
}

function insertUpTransactionsForAccount_() {
  insert_(
    '=UP_TRANSACTIONS_FOR_ACCOUNT(Yeah)',
    UP_TRANSACTIONS_HEADINGS.length,
  );
}

function insertUpAccounts_() {
  insert_('=UP_ACCOUNTS(Yeah)', UP_ACCOUNTS_HEADINGS.length);
}

function insertUpCategories_() {
  insert_('=UP_CATEGORIES(Yeah)', UP_CATEGORIES_HEADINGS.length);
}

const UP_TRANSACTIONS_HEADINGS = [
  'Created At',
  'Settled At',
  'Status',
  'Direction',
  'Currency',
  'Value',
  'Description',
  'Category',
  'Parent Category',
  'Tags',
  'Message',
];

/**
 * Retrieve transactions across all of your Up accounts.
 *
 * @param yeah Dependencies.
 * @param {string} filterQuery The filter querystring to use, e.g. `"filter[status]=HELD&filter[category]=booze"`.
 * @param {"DEBIT" | "CREDIT"} type Further filter transactions by direction (ALL/CREDIT/DEBIT).
 * @return Up Transactions
 * @example =UP_TRANSACTIONS("filter[category]=takeaway", "DEBIT") // All outgoing transactions classified as "takeaway".
 * @customfunction
 */
function UP_TRANSACTIONS(yeah, filterQuery = '', type = null) {
  return up_(`transactions?${hackyUriEncode_(filterQuery)}`, {
    tabulate: (data) => tabulateTransactions_(type, data),
  });
}

/**
 * Retrieve all transactions between two dates.
 *
 * @param yeah Dependencies.
 * @param {Date} since The start date.
 * @param {Date} until The end date.
 * @param {string} filterQuery The filter querystring to use, e.g. `"filter[status]=HELD&filter[category]=booze"`.
 * @param {"DEBIT" | "CREDIT"} type Further filter transactions by direction (ALL/CREDIT/DEBIT).
 * @example =UP_TRANSACTIONS_BETWEEN(TODAY() - 7, TODAY()) // All transactions in the last week.
 * @example =UP_TRANSACTIONS_BETWEEN(A1, B1) // All transactions between the dates set in cells `A1` and `B1`.
 * @return Up Transactions
 * @customfunction
 */
function UP_TRANSACTIONS_BETWEEN(
  yeah,
  since,
  until,
  filterQuery = '',
  type = null,
) {
  return up_(
    'transactions' +
      `?filter[since]=${encodeDate_(since)}` +
      `&filter[until]=${encodeDate_(until)}` +
      `&${hackyUriEncode_(filterQuery)}`,
    {
      tabulate: (data) => tabulateTransactions_(type, data),
    },
  );
}

/**
 * Retrieve transactions from a specific Up account.
 *
 * @param yeah Dependencies.
 * @param {string} accountId The Up Account ID.
 * @param {string} filterQuery The filter querystring to use, e.g. `"filter[status]=HELD&filter[category]=booze"`.
 * @param {"DEBIT" | "CREDIT"} type Further filter transactions by direction (ALL/CREDIT/DEBIT).
 * @return Up Transactions
 * @example =UP_TRANSACTIONS_FOR_ACCOUNT("aaaa-bbbb-cccc-dddd-eee") // All transactions for the specified account.
 * @customfunction
 */
function UP_TRANSACTIONS_FOR_ACCOUNT(
  yeah,
  accountId,
  filterQuery = '',
  type = null,
) {
  if (!accountId) {
    return 'accountId is required.';
  }

  return up_(
    `accounts/${accountId}/transactions?${hackyUriEncode_(filterQuery)}`,
    {
      tabulate: (data) => tabulateTransactions_(type, data),
    },
  );
}

function tabulateTransactions_(type, transactions) {
  if (type === 'DEBIT') {
    transactions = transactions.filter(
      (tx) => tx.attributes.amount.valueInBaseUnits < 0,
    );
  }

  if (type === 'CREDIT') {
    transactions = transactions.filter(
      (tx) => tx.attributes.amount.valueInBaseUnits > 0,
    );
  }

  const table = transactions.map((transaction) => {
    const attributes = transaction.attributes;
    return [
      new Date(attributes.createdAt),
      attributes.settledAt ? new Date(attributes.settledAt) : '',
      attributes.status,
      attributes.amount.valueInBaseUnits < 0 ? 'DEBIT' : 'CREDIT',
      attributes.amount.currencyCode,
      Math.abs(Number(attributes.amount.value)),
      attributes.description,
      transaction.relationships.category.data
        ? transaction.relationships.category.data.id
        : '',
      transaction.relationships.parentCategory.data
        ? transaction.relationships.parentCategory.data.id
        : '',
      transaction.relationships.tags.data.map((tag) => tag.id).join(','),
      attributes.message,
    ];
  });
  return [UP_TRANSACTIONS_HEADINGS, ...table];
}

const UP_ACCOUNTS_HEADINGS = [
  'Account ID',
  'Created At',
  'Type',
  'Name',
  'Currency',
  'Balance',
];

/**
 * Retrieve all your Up accounts, including balances.
 *
 * @return Up Accounts
 * @param yeah Dependencies.
 * @example =UP_ACCOUNTS() // Get all accounts.
 * @customfunction
 */
function UP_ACCOUNTS(yeah) {
  return up_('accounts', {
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
  'Category ID',
  'Category Name',
  'Parent Category ID',
];

/**
 * Retrieve all Up pre-defined categories, including parent categories.
 *
 * @return Up Categories
 * @param yeah Dependencies.
 * @example =UP_CATEGORIES() // Get all categories.
 * @customfunction
 */
function UP_CATEGORIES(yeah) {
  return up_('categories', {
    tabulate(data) {
      const table = data.map((category) => [
        category.id,
        category.attributes.name,
        category.relationships.parent.data
          ? category.relationships.parent.data.id
          : 'all',
      ]);
      return [UP_CATEGORIES_HEADINGS, ...table, ['all', 'All', '']];
    },
  });
}

const UP_TAGS_HEADINGS = ['Tag'];

/**
 * Retrieve all your user-defined tags.
 *
 * @return Up Tags
 * @param yeah Dependencies.
 * @example =UP_TAGS() // Get all tags.
 * @customfunction
 */
function UP_TAGS(yeah) {
  return up_('tags', {
    tabulate(data) {
      const table = data.map((tag) => [tag.id]);
      return [UP_TAGS_HEADINGS, ...table];
    },
  });
}

/**
 * Ping the Up API to validate your token.
 *
 * @return Up Ping
 * @example =UP_PING() // Ping the API.
 * @customfunction
 */
function UP_PING() {
  const {tokenExpiry} = TokenCache.retrieve();
  return up_('util/ping', {
    paginate: false,
    tabulate: (response) => [
      ['Up API Status', 'Token Expiry'],
      [response.meta.statusEmoji, tokenExpiry],
    ],
  });
}

function up_(path, {paginate = true, tabulate}) {
  const {token} = TokenCache.retrieve();
  if (!token) {
    return [
      'ERROR',
      'Token not provided',
      'Please navigate to "Add-ons" → "Up API" → "Set Up..."',
    ];
  }

  try {
    let url = `https://api.up.com.au/api/v1/${path}`;
    const data = [];
    do {
      const json = UrlFetchApp.fetch(url, {
        headers: {Authorization: `Bearer ${token}`},
        muteHttpExceptions: true,
      }).getContentText();
      const response = JSON.parse(json);
      if (response.errors) {
        return [['API Error']].concat(
          response.errors.map((error) => [
            error.status,
            error.title,
            error.detail,
          ]),
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
    return ['ERROR', error.message];
  }
}

function encodeDate_(date) {
  return encodeURIComponent(new Date(date).toISOString());
}

/* 🙈 */
function hackyUriEncode_(query) {
  return query
    .split('&')
    .map((kv) => {
      const [k, v] = kv.split('=');
      return `${k}=${encodeURIComponent(v)}`;
    })
    .join('&');
}
