# up-api-google-sheets

Prototype for using the [Up Banking API](https://developer.up.com.au/) in Google Sheets.

> **NOTE:** The Up API is in beta release for personal use only.

## Installation

1. Create a new Google Sheet and select "Extensions" → "Apps Script".
2. Paste in the contents of [`up.js`](https://github.com/azz/up-api-google-sheets/blob/master/up.js) into the code window.
3. Save the script, and rename the project to "Up API" and accept the authorization request (you may need to proceed through a security warning, as this is not a published Addon).
4. Click the "Run" button with the "onOpen" function selected.
5. Go back to your Google Sheet and you should have a new "Up API" drop-down. Select"Up API" → "Set Up...".

## Usage

Once you have authenticated ("Up API" → "Set Up..."), you will be able to insert formulas.

All the formulas provided by this script can be inserted from the "Up API" menu. Doing so will insert the formula, apply some styling, then auto-size the columns. Alternatively, you can enter in the formulas manually.

You can share a sheet with someone else and they will have to use their own personal access token. None of your data will be shared.

Your token will only be stored for one day. After this time your data will be cleared and you will have to re-authenticate.

## Functions

<dl>
<dt><a href="#UP_TRANSACTIONS">UP_TRANSACTIONS(yeah, filterQuery, type)</a> </dt>
<dd><p>Retrieve transactions across all of your Up accounts.</p>
</dd>
<dt><a href="#UP_TRANSACTIONS_BETWEEN">UP_TRANSACTIONS_BETWEEN(yeah, since, until, filterQuery, type)</a> </dt>
<dd><p>Retrieve all transactions between two dates.</p>
</dd>
<dt><a href="#UP_TRANSACTIONS_FOR_ACCOUNT">UP_TRANSACTIONS_FOR_ACCOUNT(yeah, accountId, filterQuery, type)</a> </dt>
<dd><p>Retrieve transactions from a specific Up account.</p>
</dd>
<dt><a href="#UP_ACCOUNTS">UP_ACCOUNTS(yeah)</a> </dt>
<dd><p>Retrieve all your Up accounts, including balances.</p>
</dd>
<dt><a href="#UP_CATEGORIES">UP_CATEGORIES(yeah)</a> </dt>
<dd><p>Retrieve all Up pre-defined categories, including parent categories.</p>
</dd>
<dt><a href="#UP_TAGS">UP_TAGS(yeah)</a> </dt>
<dd><p>Retrieve all your user-defined tags.</p>
</dd>
<dt><a href="#UP_PING">UP_PING(yeah)</a> </dt>
<dd><p>Ping the Up API to validate your token.</p>
</dd>
</dl>

<a name="UP_TRANSACTIONS"></a>

## UP_TRANSACTIONS(yeah, filterQuery, type)

Retrieve transactions across all of your Up accounts.

**Kind**: global function  
**Returns**: Up Transactions  
**Customfunction**:

| Param       | Type                                                              | Default       | Description                                                                         |
| ----------- | ----------------------------------------------------------------- | ------------- | ----------------------------------------------------------------------------------- |
| yeah        |                                                                   |               | Dependencies.                                                                       |
| filterQuery | <code>string</code>                                               |               | The filter querystring to use, e.g. `"filter[status]=HELD&filter[category]=booze"`. |
| type        | <code>&quot;DEBIT&quot;</code> \| <code>&quot;CREDIT&quot;</code> | <code></code> | Further filter transactions by direction (ALL/CREDIT/DEBIT).                        |

**Example**

```js
=UP_TRANSACTIONS(Yeah, "filter[category]=takeaway", "DEBIT") // All outgoing transactions classified as "takeaway".
```

<a name="UP_TRANSACTIONS_BETWEEN"></a>

## UP_TRANSACTIONS_BETWEEN(yeah, since, until, filterQuery, type)

Retrieve all transactions between two dates.

**Kind**: global function  
**Returns**: Up Transactions  
**Customfunction**:

| Param       | Type                                                              | Default       | Description                                                                         |
| ----------- | ----------------------------------------------------------------- | ------------- | ----------------------------------------------------------------------------------- |
| yeah        |                                                                   |               | Dependencies.                                                                       |
| since       | <code>Date</code>                                                 |               | The start date.                                                                     |
| until       | <code>Date</code>                                                 |               | The end date.                                                                       |
| filterQuery | <code>string</code>                                               |               | The filter querystring to use, e.g. `"filter[status]=HELD&filter[category]=booze"`. |
| type        | <code>&quot;DEBIT&quot;</code> \| <code>&quot;CREDIT&quot;</code> | <code></code> | Further filter transactions by direction (ALL/CREDIT/DEBIT).                        |

**Example**

```js
=UP_TRANSACTIONS_BETWEEN(Yeah, TODAY() - 7, TODAY()) // All transactions in the last week.
```

**Example**

```js
=UP_TRANSACTIONS_BETWEEN(Yeah, A1, B1) // All transactions between the dates set in cells `A1` and `B1`.
```

<a name="UP_TRANSACTIONS_FOR_ACCOUNT"></a>

## UP_TRANSACTIONS_FOR_ACCOUNT(yeah, accountId, filterQuery, type)

Retrieve transactions from a specific Up account.

**Kind**: global function  
**Returns**: Up Transactions  
**Customfunction**:

| Param       | Type                                                              | Default       | Description                                                                         |
| ----------- | ----------------------------------------------------------------- | ------------- | ----------------------------------------------------------------------------------- |
| yeah        |                                                                   |               | Dependencies.                                                                       |
| accountId   | <code>string</code>                                               |               | The Up Account ID.                                                                  |
| filterQuery | <code>string</code>                                               |               | The filter querystring to use, e.g. `"filter[status]=HELD&filter[category]=booze"`. |
| type        | <code>&quot;DEBIT&quot;</code> \| <code>&quot;CREDIT&quot;</code> | <code></code> | Further filter transactions by direction (ALL/CREDIT/DEBIT).                        |

**Example**

```js
=UP_TRANSACTIONS_FOR_ACCOUNT(Yeah, "aaaa-bbbb-cccc-dddd-eee") // All transactions for the specified account.
```

<a name="UP_ACCOUNTS"></a>

## UP_ACCOUNTS(yeah)

Retrieve all your Up accounts, including balances.

**Kind**: global function  
**Returns**: Up Accounts  
**Customfunction**:

| Param | Description   |
| ----- | ------------- |
| yeah  | Dependencies. |

**Example**

```js
=UP_ACCOUNTS(Yeah) // Get all accounts.
```

<a name="UP_CATEGORIES"></a>

## UP_CATEGORIES(yeah)

Retrieve all Up pre-defined categories, including parent categories.

**Kind**: global function  
**Returns**: Up Categories  
**Customfunction**:

| Param | Description   |
| ----- | ------------- |
| yeah  | Dependencies. |

**Example**

```js
=UP_CATEGORIES(Yeah) // Get all categories.
```

<a name="UP_TAGS"></a>

## UP_TAGS(yeah)

Retrieve all your user-defined tags.

**Kind**: global function  
**Returns**: Up Tags  
**Customfunction**:

| Param | Description   |
| ----- | ------------- |
| yeah  | Dependencies. |

**Example**

```js
=UP_TAGS(Yeah) // Get all tags.
```

<a name="UP_PING"></a>

## UP_PING(yeah)

Ping the Up API to validate your token.

**Kind**: global function  
**Returns**: Up Ping  
**Customfunction**:  
**Example**

```js
=UP_PING(Yeah) // Ping the API.
```
