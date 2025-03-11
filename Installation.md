# Installation & Setup Instructions

Follow these steps to install and run the **Google Ads Impression Share Aggregation Script**:

---

## 1. Prepare a Google Spreadsheet (Optional)

1. **Decide whether to create a new spreadsheet** manually or let the script do it automatically.
   - If you already have a dedicated spreadsheet you want to use, open it and copy its URL.
   - If you want the script to create a new spreadsheet each time, leave the `spreadsheetUrl` config field blank.

2. Make sure you have permission to edit the spreadsheet if you plan to specify a pre-existing one.

---

## 2. Add the Script in Your Google Ads Account

1. Log in to your **Google Ads** account.
2. Navigate to **Tools & Settings > Bulk Actions > Scripts**.
3. Click the **+** button to create a new script.
4. **Copy and paste** the entire script from this repository into the script editor.

---

## 3. Configure the Script

Within the script, you’ll find a `CONFIG` object at the top:

```js
const CONFIG = {
  dateRange: {
    useCustomDateRange: true,
    customDateRange: {
      start: '20250213',
      end: '20250217'
    },
    predefinedRange: 'LAST_30_DAYS'
  },
  output: {
    spreadsheetUrl: '', // Provide a spreadsheet URL or leave blank
    emailResults: true,
    emailAddress: 'your-email@example.com'
  },
  filters: {
    includeCampaignNameContains: [],
    excludeCampaignNameContains: [],
    campaignTypes: []
  },
  display: {
    showLessThan10AsText: true,
    roundPercentages: false,
    decimalPlaces: 2
  }
};
