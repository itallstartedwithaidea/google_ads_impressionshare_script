/**
 * Google Ads Impression Share Aggregation Script
 * 
 * This script properly aggregates impression share metrics using the same
 * weighting methodology as the Google Ads interface across all campaign types.
 * 
 * Metrics handled:
 * - Search Impression Share
 * - Click Share
 * - Search Exact Match Impression Share
 * - Absolute Top Impression Share
 * - Top Impression Share
 * 
 * @version 2.0
 */

// Configuration
const CONFIG = {
  // Date range options
  dateRange: {
    // Set to true to use a custom date range, false to use a predefined range
    useCustomDateRange: true,
    // Custom date range (only used if useCustomDateRange is true)
    customDateRange: {
      start: '20250213', // February 13, 2025 (YYYYMMDD format)
      end: '20250217'    // February 17, 2025 (YYYYMMDD format)
    },
    // Predefined date range options (only used if useCustomDateRange is false)
    predefinedRange: 'LAST_30_DAYS' 
  },
  
  // Output options
  output: {
    spreadsheetUrl: '', // Leave blank to create a new spreadsheet
    emailResults: true,
    emailAddress: 'your-email@example.com'
  },
  
  // Campaign filters (optional)
  filters: {
    includeCampaignNameContains: [],
    excludeCampaignNameContains: [],
    campaignTypes: [] // Leave empty for all types, or specify: 'SEARCH', 'SHOPPING', 'DISPLAY', 'VIDEO', 'PERFORMANCE_MAX'
  },
  
  // Display options
  display: {
    showLessThan10AsText: true, // Show "< 10%" instead of the actual value when below 10%
    roundPercentages: false,     // Round percentages to whole numbers
    decimalPlaces: 2             // Number of decimal places to show for percentages
  }
};

/**
 * Main function that runs the impression share aggregation
 */
function main() {
  Logger.log('Starting Impression Share Aggregation Script...');
  
  // Determine the date range to use
  const dateRange = getDateRange();
  Logger.log(`Using date range: ${dateRange.start} to ${dateRange.end}`);
  
  // Get campaign data with impression share metrics
  const campaignData = getCampaignData(dateRange);
  
  // Calculate properly weighted impression share metrics by campaign type
  const aggregatedData = calculateWeightedMetricsByCampaignType(campaignData);
  
  // Output the results
  outputResults(campaignData, aggregatedData, dateRange);
  
  Logger.log('Script completed successfully.');
}

function getDateRange() {
  if (CONFIG.dateRange.useCustomDateRange) {
    if (CONFIG.dateRange.customDateRange.start && CONFIG.dateRange.customDateRange.end) {
      return {
        start: CONFIG.dateRange.customDateRange.start,
        end: CONFIG.dateRange.customDateRange.end
      };
    }
  }
  return getDateRangeFromPredefined(CONFIG.dateRange.predefinedRange);
}

function getDateRangeFromPredefined(predefinedRange) {
  const now = new Date();
  let startDate, endDate;
  
  switch (predefinedRange) {
    case 'TODAY':
      startDate = new Date(now);
      endDate = new Date(now);
      break;
    case 'YESTERDAY':
      startDate = new Date(now);
      startDate.setDate(startDate.getDate() - 1);
      endDate = new Date(startDate);
      break;
    case 'LAST_7_DAYS':
      endDate = new Date(now);
      startDate = new Date(now);
      startDate.setDate(startDate.getDate() - 7);
      break;
    case 'LAST_WEEK':
      endDate = new Date(now);
      endDate.setDate(endDate.getDate() - (endDate.getDay() + 1));
      startDate = new Date(endDate);
      startDate.setDate(startDate.getDate() - 6);
      break;
    case 'LAST_14_DAYS':
      endDate = new Date(now);
      startDate = new Date(now);
      startDate.setDate(startDate.getDate() - 14);
      break;
    case 'LAST_30_DAYS':
      endDate = new Date(now);
      startDate = new Date(now);
      startDate.setDate(startDate.getDate() - 30);
      break;
    case 'THIS_MONTH':
      startDate = new Date(now.getFullYear(), now.getMonth(), 1);
      endDate = new Date(now);
      break;
    case 'LAST_MONTH':
      startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      endDate = new Date(now.getFullYear(), now.getMonth(), 0);
      break;
    case 'ALL_TIME':
      startDate = new Date(2000, 0, 1);
      endDate = new Date(now);
      break;
    default:
      endDate = new Date(now);
      startDate = new Date(now);
      startDate.setDate(startDate.getDate() - 30);
  }
  
  return {
    start: formatDateYYYYMMDD(startDate),
    end: formatDateYYYYMMDD(endDate)
  };
}

function formatDateYYYYMMDD(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}${month}${day}`;
}

function getCampaignData(dateRange) {
  let conditions = [
    `metrics.impressions > 0`,
    `segments.date BETWEEN ${dateRange.start} AND ${dateRange.end}`
  ];
  
  if (CONFIG.filters.campaignTypes && CONFIG.filters.campaignTypes.length > 0) {
    conditions.push(`campaign.advertising_channel_type IN (${CONFIG.filters.campaignTypes.map(t => `'${t}'`).join(',')})`);
  }
  
  const query = `
    SELECT
      campaign.id,
      campaign.name,
      campaign.status,
      campaign.advertising_channel_type,
      campaign.bidding_strategy_type,
      metrics.impressions,
      metrics.clicks,
      metrics.all_conversions,
      metrics.cost_micros,
      metrics.conversions,
      metrics.search_impression_share,
      metrics.search_top_impression_share,
      metrics.search_absolute_top_impression_share,
      metrics.search_click_share,
      metrics.search_exact_match_impression_share,
      metrics.video_view_rate,
      metrics.video_views,
      metrics.engagements
    FROM campaign
    WHERE ${conditions.join(' AND ')}
    ORDER BY metrics.impressions DESC
  `;
  
  try {
    const report = AdsApp.report(query);
    const rows = report.rows();
    const campaigns = [];
    
    while (rows.hasNext()) {
      try {
        const row = rows.next();
        
        if (CONFIG.filters.includeCampaignNameContains.length > 0) {
          const nameMatches = CONFIG.filters.includeCampaignNameContains.some(
            term => row['campaign.name'].indexOf(term) !== -1
          );
          if (!nameMatches) continue;
        }
        
        if (CONFIG.filters.excludeCampaignNameContains.length > 0) {
          const nameMatches = CONFIG.filters.excludeCampaignNameContains.some(
            term => row['campaign.name'].indexOf(term) !== -1
          );
          if (nameMatches) continue;
        }
        
        const campaignType = row['campaign.advertising_channel_type'];
        
        const searchIS = parseFloat(row['metrics.search_impression_share'] || 0) * 100;
        const clickShare = parseFloat(row['metrics.search_click_share'] || 0) * 100;
        const exactMatchIS = parseFloat(row['metrics.search_exact_match_impression_share'] || 0) * 100;
        const topIS = parseFloat(row['metrics.search_top_impression_share'] || 0) * 100;
        const absTopIS = parseFloat(row['metrics.search_absolute_top_impression_share'] || 0) * 100;
        
        const impressions = parseInt(row['metrics.impressions'] || 0, 10);
        const clicks = parseInt(row['metrics.clicks'] || 0, 10);
        const cost = parseInt(row['metrics.cost_micros'] || 0, 10) / 1000000;
        const conversions = parseFloat(row['metrics.conversions'] || 0);
        const allConversions = parseFloat(row['metrics.all_conversions'] || 0);
        
        const videoViews = parseInt(row['metrics.video_views'] || 0, 10);
        const videoViewRate = parseFloat(row['metrics.video_view_rate'] || 0) * 100;
        const engagements = parseInt(row['metrics.engagements'] || 0, 10);
        
        let interactions = clicks;
        if (campaignType === 'VIDEO') {
          interactions = engagements;
        } else if (campaignType === 'PERFORMANCE_MAX') {
          interactions = clicks + videoViews;
        }
        
        const eligibleImpressions = searchIS > 0 ? Math.round(impressions / (searchIS / 100)) : 0;
        const eligibleExactMatchImpressions = exactMatchIS > 0 ? Math.round(impressions / (exactMatchIS / 100)) : 0;
        const eligibleClicks = clickShare > 0 ? Math.round(clicks / (clickShare / 100)) : 0;
        
        campaigns.push({
          id: row['campaign.id'],
          name: row['campaign.name'],
          status: row['campaign.status'],
          type: campaignType,
          biddingStrategy: row['campaign.bidding_strategy_type'],
          impressions: impressions,
          clicks: clicks,
          interactions: interactions,
          cost: cost,
          conversions: conversions,
          allConversions: allConversions,
          searchIS: searchIS,
          clickShare: clickShare,
          exactMatchIS: exactMatchIS,
          topIS: topIS,
          absTopIS: absTopIS,
          videoViews: videoViews,
          videoViewRate: videoViewRate,
          engagements: engagements,
          eligibleImpressions: eligibleImpressions,
          eligibleExactMatchImpressions: eligibleExactMatchImpressions,
          eligibleClicks: eligibleClicks,
          hasImpressionShareMetrics: ['SEARCH', 'SHOPPING', 'PERFORMANCE_MAX'].includes(campaignType)
        });
      } catch (rowError) {
        Logger.log(`Error processing row: ${rowError}`);
      }
    }
    
    Logger.log(`Retrieved data for ${campaigns.length} campaigns`);
    return campaigns;
  } catch (reportError) {
    Logger.log(`Error running report: ${reportError}`);
    return [];
  }
}

function calculateWeightedMetricsByCampaignType(campaignData) {
  const campaignsByType = {};
  campaignData.forEach(campaign => {
    if (!campaignsByType[campaign.type]) {
      campaignsByType[campaign.type] = [];
    }
    campaignsByType[campaign.type].push(campaign);
  });
  
  const metricsByType = {};
  Object.keys(campaignsByType).forEach(type => {
    metricsByType[type] = calculateWeightedMetrics(campaignsByType[type]);
  });
  
  const overallMetrics = calculateWeightedMetrics(campaignData);
  return {
    overall: overallMetrics,
    byType: metricsByType
  };
}

function calculateWeightedMetrics(campaigns) {
  let totalImpressions = 0;
  let totalClicks = 0;
  let totalInteractions = 0;
  let totalCost = 0;
  let totalConversions = 0;
  let totalAllConversions = 0;
  let totalEligibleImpressions = 0;
  let totalEligibleExactMatchImpressions = 0;
  let totalEligibleClicks = 0;
  let totalVideoViews = 0;
  let totalEngagements = 0;
  let campaignsWithImpressionShareMetrics = 0;
  
  campaigns.forEach(campaign => {
    totalImpressions += campaign.impressions;
    totalClicks += campaign.clicks;
    totalInteractions += campaign.interactions;
    totalCost += campaign.cost;
    totalConversions += campaign.conversions;
    totalAllConversions += campaign.allConversions;
    totalVideoViews += campaign.videoViews;
    totalEngagements += campaign.engagements;
    
    if (campaign.hasImpressionShareMetrics) {
      totalEligibleImpressions += campaign.eligibleImpressions;
      totalEligibleExactMatchImpressions += campaign.eligibleExactMatchImpressions;
      totalEligibleClicks += campaign.eligibleClicks;
      campaignsWithImpressionShareMetrics++;
    }
  });
  
  let aggregatedSearchIS = 0;
  let aggregatedExactMatchIS = 0;
  let aggregatedClickShare = 0;
  
  if (campaignsWithImpressionShareMetrics > 0) {
    aggregatedSearchIS = totalEligibleImpressions > 0 ? 
      (totalImpressions / totalEligibleImpressions * 100) : 0;
    aggregatedExactMatchIS = totalEligibleExactMatchImpressions > 0 ? 
      (totalImpressions / totalEligibleExactMatchImpressions * 100) : 0;
    aggregatedClickShare = totalEligibleClicks > 0 ? 
      (totalClicks / totalEligibleClicks * 100) : 0;
  }
  
  const formattedSearchIS = formatImpressionShareMetric(aggregatedSearchIS);
  const formattedClickShare = formatImpressionShareMetric(aggregatedClickShare);
  const formattedExactMatchIS = formatImpressionShareMetric(aggregatedExactMatchIS);
  
  const ctr = totalImpressions > 0 ? (totalClicks / totalImpressions * 100) : 0;
  const convRate = totalClicks > 0 ? (totalConversions / totalClicks * 100) : 0;
  const avgCpc = totalClicks > 0 ? (totalCost / totalClicks) : 0;
  const costPerConv = totalConversions > 0 ? (totalCost / totalConversions) : 0;
  
  return {
    impressions: totalImpressions,
    clicks: totalClicks,
    interactions: totalInteractions,
    cost: totalCost.toFixed(2),
    conversions: totalConversions.toFixed(2),
    allConversions: totalAllConversions.toFixed(2),
    ctr: ctr.toFixed(2) + '%',
    convRate: convRate.toFixed(2) + '%',
    avgCpc: avgCpc.toFixed(2),
    costPerConv: costPerConv.toFixed(2),
    searchImpressionShare: formattedSearchIS,
    clickShare: formattedClickShare,
    exactMatchImpressionShare: formattedExactMatchIS,
    eligibleImpressions: totalEligibleImpressions,
    eligibleExactMatchImpressions: totalEligibleExactMatchImpressions,
    eligibleClicks: totalEligibleClicks,
    videoViews: totalVideoViews,
    engagements: totalEngagements,
    hasImpressionShareMetrics: campaignsWithImpressionShareMetrics > 0
  };
}

function formatImpressionShareMetric(value) {
  if (value === 0) {
    return '—';
  }
  
  if (value < 10) {
    return '< 10%';
  }
  
  if (CONFIG.display.roundPercentages) {
    return Math.round(value) + '%';
  }
  
  return value.toFixed(CONFIG.display.decimalPlaces) + '%';
}

function outputResults(campaignData, aggregatedData, dateRange) {
  let spreadsheet;
  if (CONFIG.output.spreadsheetUrl) {
    spreadsheet = SpreadsheetApp.openByUrl(CONFIG.output.spreadsheetUrl);
  } else {
    const formattedStartDate = dateRange.start.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3');
    const formattedEndDate = dateRange.end.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3');
    spreadsheet = SpreadsheetApp.create(`Google Ads Impression Share Report (${formattedStartDate} to ${formattedEndDate})`);
    Logger.log('Created new spreadsheet: ' + spreadsheet.getUrl());
  }
  
  const sheet = spreadsheet.getActiveSheet();
  sheet.clear();
  
  const hasVideoOrPMax = campaignData.some(c => c.type === 'VIDEO' || c.type === 'PERFORMANCE_MAX');
  
  const headers = [
    'Campaign', 'Status', 'Type', 'Bidding Strategy', 
    'Impressions', hasVideoOrPMax ? 'Interactions' : 'Clicks', 'Cost', 'Conv.', 
    'CTR', 'Conv. Rate', 'Avg. CPC', 'Cost/Conv.',
    'Search Impr. Share', 'Click Share', 'Search Exact Match IS'
  ];
  
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  campaignData.forEach(campaign => {
    const ctr = campaign.impressions > 0 ? (campaign.clicks / campaign.impressions * 100).toFixed(2) + '%' : '0%';
    const convRate = campaign.clicks > 0 ? (campaign.conversions / campaign.clicks * 100).toFixed(2) + '%' : '0%';
    const avgCpc = campaign.clicks > 0 ? (campaign.cost / campaign.clicks).toFixed(2) : '0';
    const costPerConv = campaign.conversions > 0 ? (campaign.cost / campaign.conversions).toFixed(2) : '0';
    
    const searchISFormatted = campaign.hasImpressionShareMetrics ? 
      formatImpressionShareMetric(campaign.searchIS) : '—';
    const clickShareFormatted = campaign.hasImpressionShareMetrics ? 
      formatImpressionShareMetric(campaign.clickShare) : '—';
    const exactMatchISFormatted = campaign.hasImpressionShareMetrics ? 
      formatImpressionShareMetric(campaign.exactMatchIS) : '—';
    
    sheet.appendRow([
      campaign.name,
      campaign.status,
      campaign.type,
      campaign.biddingStrategy,
      campaign.impressions,
      hasVideoOrPMax ? campaign.interactions : campaign.clicks,
      campaign.cost.toFixed(2),
      campaign.conversions.toFixed(2),
      ctr,
      convRate,
      avgCpc,
      costPerConv,
      searchISFormatted,
      clickShareFormatted,
      exactMatchISFormatted
    ]);
  });
  
  const separatorRow = Array(headers.length).fill('');
  sheet.appendRow(separatorRow);
  
  sheet.appendRow([
    'TOTAL',
    '',
    '',
    '',
    aggregatedData.overall.impressions,
    hasVideoOrPMax ? aggregatedData.overall.interactions : aggregatedData.overall.clicks,
    aggregatedData.overall.cost,
    aggregatedData.overall.conversions,
    aggregatedData.overall.ctr,
    aggregatedData.overall.convRate,
    aggregatedData.overall.avgCpc,
    aggregatedData.overall.costPerConv,
    aggregatedData.overall.searchImpressionShare,
    aggregatedData.overall.clickShare,
    aggregatedData.overall.exactMatchImpressionShare
  ]);
  
  sheet.autoResizeColumns(1, headers.length);
  
  if (CONFIG.output.emailResults) {
    const subject = 'Google Ads Impression Share Report - ' + 
      dateRange.start.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3') + ' to ' + 
      dateRange.end.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3');
    
    const body = `
      Google Ads Impression Share Report
      
      Date Range: ${dateRange.start.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3')} to ${dateRange.end.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3')}
      
      ACCOUNT TOTALS:
      Impressions: ${aggregatedData.overall.impressions}
      Clicks: ${aggregatedData.overall.clicks}
      Cost: $${aggregatedData.overall.cost}
      Conversions: ${aggregatedData.overall.conversions}
      
      Search Impression Share: ${aggregatedData.overall.searchImpressionShare}
      Click Share: ${aggregatedData.overall.clickShare}
      Search Exact Match IS: ${aggregatedData.overall.exactMatchImpressionShare}
      
      Full report available at: ${spreadsheet.getUrl()}
    `;
    
    MailApp.sendEmail({
      to: CONFIG.output.emailAddress,
      subject: subject,
      body: body
    });
    
    Logger.log(`Email report sent to: ${CONFIG.output.emailAddress}`);
  }
} 
