/**********************************************************************************
 *                                                                                *
 *   gFunctionFinance - A collection of Google App Script Functions in Finance    *
 *                                                                                *
 * ********************************************************************************
 * 
 * by LaserInvestor butterflow
 * 
 * https://blog.naver.com/bflownet
 * https://cafe.naver.com/laserinvestors
 * 
 * Version:      0.1.2
 * 
 * Copyright:    (c) 2021- by JaeWook Choi
 * 
 * License:      GNU General Public License, version 3 (GPL-3.0) 
 *               http://www.opensource.org/licenses/gpl-3.0.html
 * 
 * ********************************************************************************
 * Example google sheet                                                           *
 * ********************************************************************************
 * 
 * https://docs.google.com/spreadsheets/d/e/2PACX-1vSko6N1fc9PE0uL8CRhHKuejyMSnNEoOGB-UQyQR8lEjbWMfBzsUmEYR1B0ewPh1-GTXH2BBFAYfCWK/pubhtml
 * 
 *
 * 
 * ********************************************************************************
 * * Table of Function Scripts  *
 * ********************************************************************************
 * 
 * ADL() - Accumulation Distribution Line (https://school.stockcharts.com/doku.php?id=technical_indicators:accumulation_distribution_line)
 * ADX() - Average Directional Index (https://school.stockcharts.com/doku.php?id=technical_indicators:average_directional_index_adx)
 * Aroon() - Aroon indicator (https://school.stockcharts.com/doku.php?id=technical_indicators:aroon)
 * ATR() - Average True Range (https://school.stockcharts.com/doku.php?id=technical_indicators:average_true_range_atr)
 * Beta() - Beta to the benchmark market
 * BollingerBand() - Bollinger Bands (https://school.stockcharts.com/doku.php?id=technical_indicators:bollinger_bands)
 * CCI() - Commodity Channel Index (https://school.stockcharts.com/doku.php?id=technical_indicators:commodity_channel_index_cci)
 * ChaikinOsc() - Chaikin Oscillator (https://school.stockcharts.com/doku.php?id=technical_indicators:chaikin_oscillator)
 * Correlation() - Correlation coefficient to the benchmark market
 * EMA() - Exponential Moving Average (https://school.stockcharts.com/doku.php?id=technical_indicators:moving_averages)
 * HV() - Historical Volitality
 * MACD() - Moving Average Convergence/Divergence (https://school.stockcharts.com/doku.php?id=technical_indicators:moving_average_convergence_divergence_macd)
 * MarketCompare() - Accumulation of the difference b/w Log return of Stock A and Log return of Stock B
 * MFI() - Money Flow Index (https://school.stockcharts.com/doku.php?id=technical_indicators:money_flow_index_mfi)
 * OBV() - On Balance Volume (https://school.stockcharts.com/doku.php?id=technical_indicators:on_balance__Volume_obv)
 * PerRatio() - Stock A / Stock B
 * PivotCalculation() - add-on for PivotPoint() to calculate s3, s2, s1, r1, r2, r3 ( PivotCalculation(PivotPoint()) )
 * PivotPoint() - pivot point and highest high, lowest low and close for the previous period
 * ROC() - Rate of Change (https://school.stockcharts.com/doku.php?id=technical_indicators:rate_of_change_roc_and_momentum)
 * RRG() - Relative Rotation Graph (https://www.relativerotationgraphs.com/)
 * RSI() - Relative Strength Index (https://school.stockcharts.com/doku.php?id=technical_indicators:relative_strength_index_rsi)
 * SMA() - Simple Moving Average (https://school.stockcharts.com/doku.php?id=technical_indicators:moving_averages)
 * Stochastics() - Stochastic Oscillator (https://school.stockcharts.com/doku.php?id=technical_indicators:stochastic_oscillator_fast_slow_and_full)
 * WilliamR() - Williams %R indicator (https://school.stockcharts.com/doku.php?id=technical_indicators:williams_r)
 * WVF() - William Vix Fix (https://www.tradingview.com/script/og7JPrRA-CM-Williams-Vix-Fix-Finds-Market-Bottoms/)
 * 
 * 
 * The terms "Relative Rotation Graph" and "RRG" are registered trademarks of RRG Research(https://www.relativerotationgraphs.com/).
 * 
 * ********************************************************************************
 * Revision History                                                               *
 * ********************************************************************************
 * 0.1.0      2021.06.06  Initial release
 * 0.1.1      2021.06.15  FXCompare() added
 *                        FlipArray(), optional headerRow input argument
 *                        OptionObject(), convert array-value [] into object-key-value {}
 * 0.1.2      2021.06.24  big bug fix in ADX()
 *   
 */


/**
 * FXCompare
 * 
 * return an array with function expression comparison
 * 
 * @param {array} arrayFirst input array for the first
 * @param {number} columnFirst index for the value column (usually close price) in the first array
 * @param {array} arraySecond input array for the second
 * @param {number} columnSecond index for the value column (usually close price) in the second array
 * @param {string} functionExpression one of four predefined function expressions "add", "substract", "multiply" or "divide"  
 * @return an array with the provided function expression between first and second, [FX:funciontExpression]
 * @customfunction
 */
function FXCompare(arrayFirst, columnFirst, arraySecond, columnSecond, functionExpression) {
  // check if input arguments
  if (arguments.length < 5) throw new Error("more arguments are required");
  if (arrayFirst.map == null) throw new Error("arrayInput is not an array");
  if (arraySecond.map == null) throw new Error("arrayInput is not an array");
  if (arrayFirst.length < 2) throw new Error("arrayInput length should be greater");
  if (arrayFirst.length != arraySecond.length) throw new Error("horizontal (time serie) size of both input array should be the same");
  if (columnFirst > arrayFirst.length - 1) throw new Error("columnFirst is set out of range");
  if (columnSecond > arraySecond.length - 1) throw new Error("columnSecond is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayFirst[0].length + arraySecond[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayFirst[0];
  for (var j in arraySecond[0]) arrayResult[0].push(arraySecond[0][j]);

  // add result column headers
  arrayResult[0].push("FX:" + functionExpression.toLowerCase());

  i++;

  //
  // Stage 2: Find the first index to start ROR at
  //
  var arraySkip = 0;
  for (; i < arraySkip + 1; i++) {
    // no ROR calculation, just copy input array over
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    // skip empty data index
    if (!arrayFirst[i][columnFirst]) arraySkip++;
  }

  //
  // Stage 3: Calculate per function expression provided
  //
  for (; i < arrayFirst.length; i++) {
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    switch (functionExpression.toLowerCase()) {
      case "add":
        arrayResult[i].push(arrayFirst[i][columnFirst] + arraySecond[i][columnSecond]);
        break;

      case "substract":
        arrayResult[i].push(arrayFirst[i][columnFirst] - arraySecond[i][columnSecond]);
        break;

      case "multiply":
        arrayResult[i].push(arrayFirst[i][columnFirst] * arraySecond[i][columnSecond]);
        break;

      case "divide":
        arrayResult[i].push(arrayFirst[i][columnFirst] / arraySecond[i][columnSecond]);
        break;

      default:
        throw new Error("invalid fuction expression string is provided");
    }

  }

  return arrayResult;
}

/**
 * Relative Rotation Graph
 * 
 * return an array with RRG (The terms "Relative Rotation Graph" and "RRG" are registered trademarks of RRG Research)
 * 
 * @param {array} arrayFirst first input array for target market, size of the input array should be more than 130 for the default 12 the weekk look back period (130 = 10 x lookBackWeeks + 10)
 * @param {number} columnFirst index for the value column (usually close price) in the first array for target market
 * @param {array} arraySecond second input array for benchmark market
 * @param {number} columnSecond index for the value column (usually close price) in the second array for benchmark 
 * @param {number} lookBackWeek [OPTIONAL] 12 by default, look back period in weeks
 * @param {number} columnFirstDate index for th date column (usually column index 0) in the target market array
 * @return an array with Relative Rotation Graph, [RS, RM]
 * @customfunction
 */
function RRG(arrayFirst, columnFirst, arraySecond, columnSecond, lookBackWeek = 12, columnFirstDate = DATE) {
  // check if input arguments
  if (arguments.length < 4) throw new Error("more arguments are required");
  if (arrayFirst.map == null) throw new Error("arrayInput is not an array");
  if (arraySecond.map == null) throw new Error("arrayInput is not an array");
  if (arrayFirst.length < lookBackWeek * 10 + 10) throw new Error("arrayInput length should be greater (minimum 130 trading days)");
  if (arrayFirst.length != arraySecond.length) throw new Error("horizontal (time serie) size of both input array should be the same");
  if (columnFirst > arrayFirst.length - 1) throw new Error("columnFirst is set out of range");
  if (columnFirstDate > arrayFirst.length - 1) throw new Error("columnFirstDate is set out of range");
  if (columnSecond > arraySecond.length - 1) throw new Error("columnSecond is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayFirst[0].length + arraySecond[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayFirst[0];
  for (var j in arraySecond[0]) arrayResult[0].push(arraySecond[0][j]);

  // add result column headers
  arrayResult[0].push("RS", "RM");

  i++;

  //
  // Stage 2: Find the first index to start RRG at
  //
  var arraySkip = 0;
  for (; i < arraySkip + 1; i++) {
    // no calculation, just copy input array over
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    // skip empty data index
    if (!arrayFirst[i][columnFirst]) arraySkip++;
  }

  //
  // Stage 3: Calculate RRG
  //
  arrayResult[i] = arrayFirst[i];
  for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

  var indexWeekStart = i;

  var prevDay = arrayResult[i][columnFirstDate].getDay();
  i++

  var priceRelative = [], priceRelativeRS = [];
  var highhigh = 0, lowlow = 0, closeclose = 0, pivotPoint = 0;
  for (; i < arrayFirst.length; i++) {
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    var currentDay = arrayResult[i][columnFirstDate].getDay();

    if (currentDay < prevDay) {

      priceRelative.push(100 * arrayFirst[i][columnFirst] / arraySecond[i][columnSecond]);
      // arrayResult[i].push("new week", priceRelative.slice(-1)[0]);

      if (priceRelative.length == lookBackWeek) {

        var RS = 101 + (priceRelative.slice(-1)[0] - __VAvg(priceRelative)) / __VStdev(priceRelative);
        arrayResult[i].push(RS);
        priceRelativeRS.push(RS);
        priceRelative.shift();
      }

      if (priceRelativeRS.length == lookBackWeek + 1) {
        var RSROC = [];
        for (var j = 0; j < lookBackWeek; j++) {
          RSROC.push(100 * (priceRelativeRS[j + 1] / priceRelativeRS[j]) - 1);
        }
        arrayResult[i].push(101 + (RSROC.slice(-1)[0] - __VAvg(RSROC)) / __VStdev(RSROC));
        priceRelativeRS.shift();
      }

      indexWeekStart = i;
    }
    prevDay = currentDay;
  }

  //
  // calculate RS and RM for the last day
  //
  if (i - 1 != indexWeekStart) {
    var RS = 101 + (priceRelative.slice(-1)[0] - __VAvg(priceRelative)) / __VStdev(priceRelative);
    arrayResult[i - 1].push(RS);
    priceRelativeRS.push(RS);

    var RSROC = [];
    for (var j = 0; j < priceRelativeRS.length - 1; j++) {
      RSROC.push(100 * (priceRelativeRS[j + 1] / priceRelativeRS[j]) - 1);
    }
    arrayResult[i - 1].push(101 + (RSROC.slice(-1)[0] - __VAvg(RSROC)) / __VStdev(RSROC));
  }

  return arrayResult;
}

/**
 * PerRatio
 * 
 * return an array with ratio of first / second
 * 
 * @param {array} arrayFirst input array for the first
 * @param {number} columnFirst index for the value column (usually close price) in the first array
 * @param {array} arraySecond input array for the second
 * @param {number} columnSecond index for the value column (usually close price) in the second array
* @return an array with ratio of first / second, [PerRatio]
 * @customfunction
 */
function PerRatio(arrayFirst, columnFirst, arraySecond, columnSecond) {
  // check if input arguments
  if (arguments.length < 4) throw new Error("more arguments are required");
  if (arrayFirst.map == null) throw new Error("arrayInput is not an array");
  if (arraySecond.map == null) throw new Error("arrayInput is not an array");
  if (arrayFirst.length < 2) throw new Error("arrayInput length should be greater");
  if (arrayFirst.length != arraySecond.length) throw new Error("horizontal (time serie) size of both input array should be the same");
  if (columnFirst > arrayFirst.length - 1) throw new Error("columnFirst is set out of range");
  if (columnSecond > arraySecond.length - 1) throw new Error("columnSecond is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayFirst[0].length + arraySecond[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayFirst[0];
  for (var j in arraySecond[0]) arrayResult[0].push(arraySecond[0][j]);

  // add result column headers
  arrayResult[0].push("PerRatio");

  i++;

  //
  // Stage 2: Find the first index to start ROR at
  //
  var arraySkip = 0;
  for (; i < arraySkip + 1; i++) {
    // no ROR calculation, just copy input array over
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    // skip empty data index
    if (!arrayFirst[i][columnFirst]) arraySkip++;
  }

  //
  // Stage 3: Calculate Per Ratio of First / Second
  //
  for (; i < arrayFirst.length; i++) {
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    arrayResult[i].push(arrayFirst[i][columnFirst] / arraySecond[i][columnSecond]);
  }

  return arrayResult;
}

/**
 * MarketCompare
 * 
 * return an array with LOG return comparison
 * Return of rate is calculated by LOG return, LN(currentPrice / prevPrice), respectively
 * Accumulation of LOG return comparison, SUM( LOG return(A) - LOG return(B) ) is calculated as well
 * 
 * @param {array} arrayFirst first input array
 * @param {number} columnFirst index for the value column (usually close price) in the first array
 * @param {array} arraySecond second input array
 * @param {number} columnSecond index for the value column (usually close price) in the second array
* @return an array with LOG return comparison, [ROR1, ROR2, ROR1/ROR2]
 * @customfunction
 */
function MarketCompare(arrayFirst, columnFirst, arraySecond, columnSecond) {
  // check if input arguments
  if (arguments.length < 4) throw new Error("more arguments are required");

  if (arrayFirst.map == null) throw new Error("arrayInput is not an array");
  if (arraySecond.map == null) throw new Error("arrayInput is not an array");
  if (arrayFirst.length < 2) throw new Error("arrayInput length should be greater");
  if (arrayFirst.length != arraySecond.length) throw new Error("horizontal (time serie) size of both input array should be the same");
  if (columnFirst > arrayFirst.length - 1) throw new Error("columnFirst is set out of range");
  if (columnSecond > arraySecond.length - 1) throw new Error("columnSecond is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayFirst[0].length + arraySecond[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayFirst[0];
  for (var j in arraySecond[0]) arrayResult[0].push(arraySecond[0][j]);

  // add result column headers
  arrayResult[0].push("ROR1", "ROR2", "ROR1/ROR2");

  i++;

  //
  // Stage 2: Find the first index to start ROR at
  //
  var arraySkip = 0;
  for (; i < arraySkip + 2; i++) {
    // no ROR calculation, just copy input array over
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    // skip empty data index
    if (!arrayFirst[i][columnFirst]) arraySkip++;
  }

  //
  // Stage 3: Calculate Return of Rate
  //
  arrayResult[i] = arrayFirst[i];
  for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

  arrayResult[i].push(Math.log(arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst]));
  arrayResult[i].push(Math.log(arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond]));
  arrayResult[i].push(arrayResult[i][rIndex + 1] - arrayResult[i][rIndex + 2]);
  i++

  for (; i < arrayFirst.length; i++) {
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    arrayResult[i].push(Math.log(arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst]));
    arrayResult[i].push(Math.log(arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond]));
    arrayResult[i].push(arrayResult[i - 1][rIndex + 3] + arrayResult[i][rIndex + 1] - arrayResult[i][rIndex + 2]);
  }

  return arrayResult;
}

/**
 * Beta
 * 
 * return an array with beta calculated, Covariance(A, B) / Variance(B) 
 * 
 * @param {array} arrayFirst first input array
 * @param {number} columnFirst index for the value column (usually close price) in the first array
 * @param {array} arraySecond second input array for benchmark market
 * @param {number} columnSecond index for the value column (usually close price) in the second array
 * @param {number} lookBack [OPTIONAL] 126 by default, look back period for beta calculation, 22 for a monthly, 63 for a quarterly, 126 for semiannually, 252 for annually
 * @param {bool} logReturn [OPTIONAL] false by default, true for LOG return calculation, LN(currentPrice/prevPrice)
 * @return an array with beta, [RoR1, RoR2, beta]
 * @customfunction
 */
function Beta(arrayFirst, columnFirst, arraySecond, columnSecond, lookBack = 126, logReturn = false) {
  // check if input arguments 
  if (arguments.length < 4) throw new Error("more arguments are required");
  if (arrayFirst.map == null) throw new Error("arrayFirst is not an array");
  if (arraySecond.map == null) throw new Error("arraySecond is not an array");
  if (arrayFirst.length != arraySecond.length) throw new Error("horizontal (time serie) size of both input array should be the same");
  if (arrayFirst.length < lookBack + 2) throw new Error("arrayFirst length should be greater");
  if (columnFirst > arrayFirst[0].length - 1) throw new Error("columnFirst is set out of range");
  if (columnSecond > arraySecond[0].length - 1) throw new Error("columnSecond is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayFirst[0].length + arraySecond[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayFirst[0];
  for (var j in arraySecond[0]) arrayResult[0].push(arraySecond[0][j]);

  // add result column headers
  arrayResult[0].push("ROR1", "ROR2", "beta");
  i++;

  //
  // Stage 2: Find the first index to start beta at
  //
  var arraySkip = 0;
  for (; i < arraySkip + 2; i++) {
    // no calculation, just copy input array over
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    // skip empty data index
    if (!arrayFirst[i][columnFirst]) arraySkip++;
  }

  var lbStart = arraySkip + 2;

  //
  // Stage 3: Calculate Return of Rate
  //
  arrayResult[i] = arrayFirst[i];
  for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

  arrayResult[i].push((logReturn ? Math.log(arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst]) : (arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst] - 1)));
  arrayResult[i].push((logReturn ? Math.log(arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond]) : (arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond] - 1)));
  i++

  for (; i < lbStart + lookBack - 1; i++) {
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    arrayResult[i].push((logReturn ? Math.log(arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst]) : (arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst] - 1)));
    arrayResult[i].push((logReturn ? Math.log(arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond]) : (arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond] - 1)));

  }

  //
  // Stage 4: Calculate beta
  //
  for (; i < arrayFirst.length; i++) {
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    arrayResult[i].push((logReturn ? Math.log(arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst]) : (arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst] - 1)));
    arrayResult[i].push((logReturn ? Math.log(arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond]) : (arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond] - 1)));

    arrayResult[i].push(__HCovar(arrayResult, rIndex + 1, arrayResult, rIndex + 2, i - lookBack + 1, lookBack) / __HVar(arrayResult, rIndex + 2, i - lookBack + 1, lookBack));
  }

  return arrayResult;
}

/**
 * Correlation Coefficient
 * 
 * return Correlation coefficient, Covariance (A, B) / STDEV(A) / STDEV(B)
 * 
 * @param {array} arrayFirst first input array
 * @param {number} columnFirst index for the value column (usually close price) in the first array
 * @param {array} arraySecond second input array for benchmark market
 * @param {number} columnSecond index for the value column (usually close price) in the second array
 * @param {number} lookBack [OPTIONAL] 126 by default, look back period for correlation coefficient calculation, 22 for a monthly, 63 for a quarterly, 126 for semiannually, 252 for annually
 * @param {bool} logReturn [OPTIONAL] false by default, true for LOG return calculation, LN(currentPrice/prevPrice),
 * @return an array with Correlation Coefficient, [ROR1, ROR2, CorrCoeff]
 * @customfunction
 */
function Correlation(arrayFirst, columnFirst, arraySecond, columnSecond, lookBack = 126, logReturn = false) {
  // check if input arguments 
  if (arguments.length < 4) throw new Error("more arguments are required");
  if (arrayFirst.map == null) throw new Error("arrayFirst is not an array");
  if (arraySecond.map == null) throw new Error("arraySecond is not an array");
  if (arrayFirst.length != arraySecond.length) throw new Error("horizontal (time serie) size of both input array should be the same");
  if (arrayFirst.length < lookBack + 2) throw new Error("arrayFirst length should be greater");
  if (columnFirst > arrayFirst[0].length - 1) throw new Error("columnFirst is set out of range");
  if (columnSecond > arraySecond[0].length - 1) throw new Error("columnSecond is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayFirst[0].length + arraySecond[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayFirst[0];
  for (var j in arraySecond[0]) arrayResult[0].push(arraySecond[0][j]);

  // add result column headers
  arrayResult[0].push("ROR1", "ROR2", "CorrCoeff");
  i++;

  //
  // Stage 2: Find the first index to start Correlation coefficient at
  //
  var arraySkip = 0;
  for (; i < arraySkip + 2; i++) {
    // no calculation, just copy input array over
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    // skip empty data index
    if (!arrayFirst[i][columnFirst]) arraySkip++;
  }

  var lbStart = arraySkip + 2;

  //
  // Stage 3: Calculate Return of Rate
  //
  arrayResult[i] = arrayFirst[i];
  for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

  arrayResult[i].push((logReturn ? Math.log(arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst]) : (arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst] - 1)));
  arrayResult[i].push((logReturn ? Math.log(arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond]) : (arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond] - 1)));
  i++

  for (; i < lbStart + lookBack - 1; i++) {
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    arrayResult[i].push((logReturn ? Math.log(arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst]) : (arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst] - 1)));
    arrayResult[i].push((logReturn ? Math.log(arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond]) : (arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond] - 1)));

  }

  //
  // Stage 4: Calculate Correlation coefficient
  //
  for (; i < arrayFirst.length; i++) {
    arrayResult[i] = arrayFirst[i];
    for (var j in arraySecond[i]) arrayResult[i].push(arraySecond[i][j]);

    arrayResult[i].push((logReturn ? Math.log(arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst]) : (arrayFirst[i][columnFirst] / arrayFirst[i - 1][columnFirst] - 1)));
    arrayResult[i].push((logReturn ? Math.log(arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond]) : (arraySecond[i][columnSecond] / arraySecond[i - 1][columnSecond] - 1)));

    arrayResult[i].push(__HCovar(arrayResult, rIndex + 1, arrayResult, rIndex + 2, i - lookBack + 1, lookBack) / __HStdev(arrayResult, rIndex + 1, i - lookBack + 1, lookBack) / __HStdev(arrayResult, rIndex + 2, i - lookBack + 1, lookBack));
  }

  return arrayResult;
}

/**
 * PivotCalculation
 * 
 * return an array with Pivot Point Calculation
 * Highest of high of previous peroid, Lowest of low for previous period, close of previous peroid and pivot point are calculated
 * Use PivotPoint() function with the returned array from this function to calculate S3, S2, S1, R1, R2, R3
 * 
 * @param {array} arrayInput input array, returned array from  PivotPoint() function only
 * @return an array with all pivot points, [S3, S2, S1, R1, R2, R3]
 * @customfunction
 */
function PivotCalculation(arrayInput) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < 1) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("S3", "S2", "S1", "R1", "R2", "R3");
  i++;

  //
  // Stage 2: Find the first index to start Pivot Points at
  //
  var arraySkip = 0;
  for (; i < arraySkip + 1; i++) {
    // no calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][rIndex]) arraySkip++;
  }

  //
  // Stage 3: Calculate Pivot Points
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    if (arrayResult[i][rIndex]) {
      var highhigh = arrayResult[i][rIndex - 3];
      var lowlow = arrayResult[i][rIndex - 2];
      var closeclose = arrayResult[i][rIndex - 1];
      var pivotPoint = arrayResult[i][rIndex];

      var s1 = 2 * pivotPoint - highhigh;
      var r1 = 2 * pivotPoint - lowlow;
      var s2 = pivotPoint - (highhigh - lowlow);
      var r2 = pivotPoint + (highhigh - lowlow);
      var s3 = lowlow - 2 * (highhigh - pivotPoint);
      var r3 = highhigh + 2 * (pivotPoint - lowlow);

      arrayResult[i].push(s3, s2, s1, r1, r2, r3);
    }
  }

  return arrayResult;
}

/**
 * PivotPoint
 * 
 * return an array with Pivot Point Calculation
 * Highest of high of previous peroid, Lowest of low for previous period, close of previous peroid and pivot point are calculated
 * Use PivotPoint() function with the returned array from this function to calculate S3, S2, S1, R1, R2, R3
 *  
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} periodOption [OPTIONAL] 0 by default, 0 for daily, 1 for weekly, 2 for monthly, 3 for yearly
 * @return an array with pivot point, [PrevHigh, PrevLow, PrevClose, PivotPoint]
 * @customfunction
 */
function PivotPoint(arrayInput, periodOption = 0) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < 7) throw new Error("arrayInput length should be greater");

  const DAILY = 0, WEEKLY = 1, MONTHLY = 2, YEARLY = 3;

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("PrevHigh", "PrevLow", "PrevClose", "PivotPoint");
  i++;

  //
  // Stage 2: Find the first index to start Pivot Points at
  //
  var arraySkip = 0;
  for (; i < arraySkip + 1; i++) {
    // no calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][CLOSE]) arraySkip++;
  }

  //
  // Stage 3: Calculate Pivot Points
  //
  arrayResult[i] = arrayInput[i];

  var indexYearStart = i; indexMonthStart = i; indexWeekStart = i;

  var prevMonth = arrayResult[i][DATE].getMonth();
  var prevDate = arrayResult[i][DATE].getDate();
  var prevDay = arrayResult[i][DATE].getDay();
  i++

  var highhigh = 0, lowlow = 0, closeclose = 0, pivotPoint = 0;
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    var currentMonth = arrayResult[i][DATE].getMonth();
    var currentDate = arrayResult[i][DATE].getDate();
    var currentDay = arrayResult[i][DATE].getDay();

    switch (periodOption) {
      case WEEKLY:
        if (currentDay < prevDay) {
          highhigh = __HMax(arrayResult, HIGH, indexWeekStart, i - indexWeekStart);
          lowlow = __HMin(arrayResult, LOW, indexWeekStart, i - indexWeekStart);
          closeclose = arrayResult[i - 1][CLOSE];
          pivotPoint = (highhigh + lowlow + closeclose) / 3;

          arrayResult[i].push(highhigh, lowlow, closeclose, pivotPoint);
          indexWeekStart = i;

          delete arrayResult[i - 1][rIndex + 1];
          delete arrayResult[i - 1][rIndex + 2];
          delete arrayResult[i - 1][rIndex + 3];
          delete arrayResult[i - 1][rIndex + 4];

        } else {
          if (highhigh) {
            arrayResult[i].push(highhigh, lowlow, closeclose, pivotPoint);
          }
        }
        break;

      case MONTHLY:
        if (currentDate < prevDate) {
          highhigh = __HMax(arrayResult, HIGH, indexMonthStart, i - indexMonthStart);
          lowlow = __HMin(arrayResult, LOW, indexMonthStart, i - indexMonthStart);
          closeclose = arrayResult[i - 1][CLOSE];
          pivotPoint = (highhigh + lowlow + closeclose) / 3;

          arrayResult[i].push(highhigh, lowlow, closeclose, pivotPoint);
          indexMonthStart = i;

          delete arrayResult[i - 1][rIndex + 1];
          delete arrayResult[i - 1][rIndex + 2];
          delete arrayResult[i - 1][rIndex + 3];
          delete arrayResult[i - 1][rIndex + 4];

        } else
          if (highhigh) {
            arrayResult[i].push(highhigh, lowlow, closeclose, pivotPoint);
          }


        break;

      case YEARLY:
        if (currentMonth < prevMonth) {
          highhigh = __HMax(arrayResult, HIGH, indexYearStart, i - indexYearStart);
          lowlow = __HMin(arrayResult, LOW, indexYearStart, i - indexYearStart);
          closeclose = arrayResult[i - 1][CLOSE];
          pivotPoint = (highhigh + lowlow + closeclose) / 3;

          arrayResult[i].push(highhigh, lowlow, closeclose, pivotPoint);
          indexYearStart = i;

          delete arrayResult[i - 1][rIndex + 1];
          delete arrayResult[i - 1][rIndex + 2];
          delete arrayResult[i - 1][rIndex + 3];
          delete arrayResult[i - 1][rIndex + 4];

        } else
          if (highhigh) {
            arrayResult[i].push(highhigh, lowlow, closeclose, pivotPoint);
          }

        break;

      case DAILY:
      default:
        arrayResult[i].push(arrayResult[i - 1][HIGH], arrayResult[i - 1][LOW], arrayResult[i - 1][CLOSE]);
        indexYearStart = i;

    }

    prevMonth = currentMonth;
    prevDate = currentDate;
    prevDay = currentDay;
  }

  return arrayResult;
}

/**
 * Aroon
 * 
 * return an array with Aroon
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} lookBack [OPTIONAL] 20 by default, look back period for Aroon calculation
 * @return an array with Aroon, [Aroon]
 * @customfunction
 */
function Aroon(arrayInput, lookBack = 25) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 1) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("Aroon Up", "Aroon Down");
  i++;

  //
  // Stage 2: Find the first index to start Aroon at
  //
  var arraySkip = 0;
  for (; i < lookBack + arraySkip; i++) {
    // no Aroon calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][CLOSE]) arraySkip++;
  }

  //
  // Stage 3: Calculate Aroon
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    var lbStart = i - lookBack + 1;

    var upTrend = 0, downTrend = 0;
    var highhigh = arrayResult[lbStart][HIGH], lowlow = arrayResult[lbStart][LOW];
    for (var j = lbStart; j < lbStart + lookBack; j++) {
      if (arrayResult[j][HIGH] > highhigh) { highhigh = arrayResult[j][HIGH]; upTrend = j - lbStart + 1; }
      if (arrayResult[j][LOW] < lowlow) { lowlow = arrayResult[j][LOW]; downTrend = j - lbStart + 1; }
    }

    arrayResult[i].push(upTrend / lookBack * 100, downTrend / lookBack * 100);
  }

  return arrayResult;
}

/**
 * Williams %R
 * 
 * return an array with Williams %R
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} lookBack [OPTIONAL] 20 by default, look back period for Williams %R calcuation
 * @return an array with Williams %R, [W%R]
 * @customfunction
 */
function WilliamR(arrayInput, lookBack = 14) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 1) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("W%R");
  i++;

  //
  // Stage 2: Find the first index to start W%R at
  //
  var arraySkip = 0;
  for (; i < lookBack + arraySkip; i++) {
    // no W%R calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][CLOSE]) arraySkip++;
  }

  //
  // Stage 3: Calculate Williams %R
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    var lbStart = i - lookBack + 1;

    var highhigh = arrayResult[lbStart][HIGH], lowlow = arrayResult[lbStart][LOW];
    for (var j = lbStart; j < lbStart + lookBack; j++) {
      if (arrayResult[j][HIGH] > highhigh) highhigh = arrayResult[j][HIGH];
      if (arrayResult[j][LOW] < lowlow) lowlow = arrayResult[j][LOW];
    }

    arrayResult[i].push((highhigh - arrayResult[i][CLOSE]) / (highhigh - lowlow) * -100);
  }

  return arrayResult;
}

/**
 * CommodityChannelIndex
 * 
 * return an array with CCI
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} lookBack [OPTIONAL] 20 by default, look back period for CCI calcuation
 * @return an array with CCI, [CCI]
 * @customfunction
 */
function CCI(arrayInput, lookBack = 20, meanDeviationConstant = 0.015) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 1) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("CCI");
  i++;

  //
  // Stage 2: Find the first index to start CCI at
  //
  var arraySkip = 0;
  for (; i < lookBack + arraySkip; i++) {
    // no CCi calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][CLOSE]) arraySkip++;
  }

  //
  // Stage 3: Calculate CCI
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    var lbStart = i - lookBack + 1;

    var arrayTypicalPrice = [], typicalPriceSum = 0;
    for (var j = lbStart; j < lbStart + lookBack; j++) {
      var typicalPrice = (arrayResult[j][HIGH] + arrayResult[j][LOW] + arrayResult[j][CLOSE]) / 3;
      typicalPriceSum += typicalPrice;
      arrayTypicalPrice.push(typicalPrice);
    }

    var typicalPriceSMA = typicalPriceSum / lookBack;
    var meanDeviation = 1 * arrayTypicalPrice.map(ar => Math.abs(ar - typicalPriceSMA)).reduce((a, b) => a + b, 0) / lookBack;
    arrayResult[i].push((arrayTypicalPrice[arrayTypicalPrice.length - 1] - typicalPriceSMA) / (meanDeviationConstant * meanDeviation));
  }

  return arrayResult;
}

/**
 * MoneyFlowIndex
 * 
 * return an array with MFI
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} lookBack [OPTIONAL] 14 by default, look back period for MFI calculation
 * @return an array with MFI, [MFI]
 * @customfunction
 */
function MFI(arrayInput, lookBack = 14) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 2) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("MFI");
  i++;

  //
  // Stage 2: Find the first index to start MFI at
  //
  var arraySkip = 0;
  for (; i < lookBack + arraySkip + 1; i++) {
    // no MFI calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][CLOSE]) arraySkip++;
  }

  //
  // Stage 3: Calculate Money Flow Index
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    var lbStart = i - lookBack + 1;

    var moneyFlowPlusSum = 0, moneyFlowMinusSum = 0;
    for (var j = lbStart; j < lbStart + lookBack; j++) {
      var typicalPricePrev = (arrayResult[j - 1][HIGH] + arrayResult[j - 1][LOW] + arrayResult[j - 1][CLOSE]) / 3;
      var typicalPrice = (arrayResult[j][HIGH] + arrayResult[j][LOW] + arrayResult[j][CLOSE]) / 3;

      moneyFlowPlusSum += (typicalPrice > typicalPricePrev ? typicalPrice * arrayResult[j][VOLUME] : 0);
      moneyFlowMinusSum += (typicalPrice < typicalPricePrev ? typicalPrice * arrayResult[j][VOLUME] : 0);
    }

    arrayResult[i].push(100 - 100 / (1 + moneyFlowPlusSum / moneyFlowMinusSum));
  }

  return arrayResult;
}

/**
 * ChaikinOscillator
 * 
 * return an array with ChainkinOscillator
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @return an array with ChainkinOscillator, [ChainkinOsc]
 * @customfunction
 */
function ChaikinOsc(arrayInput) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < 2) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;
  var lookBackFast = 3, lookBackSlow = 10;
  var smoothingFast = 2 / (lookBackFast + 1), smoothingSlow = 2 / (lookBackSlow + 1);

  arrayResult = ADL(arrayInput);

  //
  // Stage 1: Header copy
  // 
  arrayResult[0].push("fastEMA", "slowEMA", "chainkinOsc");
  i++;

  //
  // Stage 2: Find the first index to start Fast EMA at
  //
  var arraySkip = 0;
  for (; i < arraySkip + lookBackFast; i++) {
    // skip empty data index
    if (!arrayInput[i][rIndex + 1]) arraySkip++;
  }

  var lbStartFast = arraySkip + 1;
  // var indexFastAt = lbStartFast + lookBackFast - 1;

  var lbStartSlow = arraySkip + 1;
  var indextSlowAt = lbStartSlow + lookBackSlow - 1;

  // the first fast average calculation with SMA

  arrayResult[i][rIndex + 2] = __HAvg(arrayResult, rIndex + 1, lbStartFast, lookBackFast);
  i++

  // fast EMA
  for (; i < indextSlowAt; i++) {
    // arrayResult[i] = arrayInput[i];

    arrayResult[i][rIndex + 2] = smoothingFast * arrayResult[i][rIndex + 1] + (1 - smoothingFast) * arrayResult[i - 1][rIndex + 2];
  }

  // the first slow average calculation with SMA 

  arrayResult[i][rIndex + 2] = smoothingFast * arrayResult[i][rIndex + 1] + (1 - smoothingFast) * arrayResult[i - 1][rIndex + 2];
  arrayResult[i][rIndex + 3] = __HAvg(arrayResult, rIndex + 1, lbStartSlow, lookBackSlow);
  arrayResult[i][rIndex + 4] = arrayResult[i][rIndex + 2] - arrayResult[i][rIndex + 3]
  i++

  // the all the rest with EMA for both fast and slow and CO = fastEMA - slowEMA
  for (; i < arrayResult.length; i++) {

    arrayResult[i][rIndex + 2] = smoothingFast * arrayResult[i][rIndex + 1] + (1 - smoothingFast) * arrayResult[i - 1][rIndex + 2];
    arrayResult[i][rIndex + 3] = smoothingSlow * arrayResult[i][rIndex + 1] + (1 - smoothingSlow) * arrayResult[i - 1][rIndex + 3];
    arrayResult[i][rIndex + 4] = arrayResult[i][rIndex + 2] - arrayResult[i][rIndex + 3]
  }

  return arrayResult;
}

/**
 * OnBalanceVolume
 * 
 * return an array with OBV
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @return an array with OBV, [OBV]
 * @customfunction
 */
function OBV(arrayInput) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < 2) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("OBV");
  i++;

  // 
  arrayResult[i] = arrayInput[i];

  arrayResult[i].push(arrayResult[i][VOLUME]);
  i++

  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    arrayResult[i].push(arrayResult[i - 1][rIndex + 1] + ((arrayResult[i - 1][CLOSE] == arrayResult[i][CLOSE]) ? 0 : Math.sign(arrayResult[i][CLOSE] - arrayResult[i - 1][CLOSE]) * arrayResult[i][VOLUME]));
  }

  return arrayResult;
}

/**
 * AccumulationDistributionLine
 * 
 * return an array with ADL
 * 
 * @param {array} arrayInput input array
 * @return an array with ADL, [ADL]
 * @customfunction
 */
function ADL(arrayInput) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < 2) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("ADL");
  i++;

  // 
  arrayResult[i] = arrayInput[i];

  var moneyFlowMult = ((arrayResult[i][CLOSE] - arrayResult[i][LOW]) - (arrayResult[i][HIGH] - arrayResult[i][CLOSE])) / (arrayResult[i][HIGH] - arrayResult[i][LOW]);
  var moneyFlowVol = moneyFlowMult * arrayResult[i][VOLUME];
  var accuDistLine = moneyFlowVol;

  arrayResult[i].push(accuDistLine);
  i++

  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    moneyFlowMult = ((arrayResult[i][CLOSE] - arrayResult[i][LOW]) - (arrayResult[i][HIGH] - arrayResult[i][CLOSE])) / (arrayResult[i][HIGH] - arrayResult[i][LOW]);
    moneyFlowVol = moneyFlowMult * arrayResult[i][VOLUME];
    accuDistLine = arrayResult[i - 1][rIndex + 1] + moneyFlowVol;

    arrayResult[i].push(accuDistLine);
  }

  return arrayResult;
}

/**
 * Stochastics
 * 
 * return an array with Stochastics
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} lookBackFast [OPTIONAL] 14 by default, look back period for fast stochastic calculation
 * @param {number} smoothingFast [OPTIONAL] 3 by default, smoothing period for fast stochastics
 * @param {number} smoothingSlow [OPTIONAL] 3 by default, smoothing period for slow stochastics
 * @return an array with stochastics, [fast%K, slow%K, slow%D]
 * @customfunction
 */
function Stochastics(arrayInput, lookBackFast = 14, smoothingFast = 3, smoothingSlow = 3) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBackFast + smoothingFast + smoothingSlow - 1) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("fask%K", "slow%K", "slow%D");
  i++;

  //
  // Stage 2: Find the first index to start Fast EMA at
  //
  var arraySkip = 0;
  for (; i < arraySkip + lookBackFast; i++) {
    // no MACD calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][CLOSE]) arraySkip++;
  }

  var lbStartFast = arraySkip + 1;
  var indexFastAt = lbStartFast + lookBackFast - 1;

  var lbStartsmoothingFast = indexFastAt;
  var indextsmoothingFastAt = lbStartsmoothingFast + smoothingFast - 1;

  var lbStartsmoothingSlow = indextsmoothingFastAt;
  var indexsmoothingSlowAt = lbStartsmoothingSlow + smoothingSlow - 1;

  var highhigh = 0, lowlow = 0;
  for (; i < indextsmoothingFastAt; i++) {
    arrayResult[i] = arrayInput[i];

    highhigh = __HMax(arrayResult, HIGH, i - lookBackFast + 1, lookBackFast);
    lowlow = __HMin(arrayResult, LOW, i - lookBackFast + 1, lookBackFast);

    // fast%K
    arrayResult[i].push((arrayResult[i][CLOSE] - lowlow) / (highhigh - lowlow));
  }

  for (; i < indexsmoothingSlowAt; i++) {
    arrayResult[i] = arrayInput[i];

    highhigh = __HMax(arrayResult, HIGH, i - lookBackFast + 1, lookBackFast);
    lowlow = __HMin(arrayResult, LOW, i - lookBackFast + 1, lookBackFast);

    // fast%K
    arrayResult[i].push((arrayResult[i][CLOSE] - lowlow) / (highhigh - lowlow));

    // slow%K
    arrayResult[i].push(__HAvg(arrayResult, rIndex + 1, i - smoothingFast + 1, smoothingFast));
  }

  // the rest calcuation
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    highhigh = __HMax(arrayResult, HIGH, i - lookBackFast + 1, lookBackFast);
    lowlow = __HMin(arrayResult, LOW, i - lookBackFast + 1, lookBackFast);

    // fast%K
    arrayResult[i].push((arrayResult[i][CLOSE] - lowlow) / (highhigh - lowlow));

    // slow%K
    arrayResult[i].push(fastD = __HAvg(arrayResult, rIndex + 1, i - smoothingFast + 1, smoothingFast));

    // slow%D
    arrayResult[i].push(__HAvg(arrayResult, rIndex + 2, i - smoothingSlow + 1, smoothingSlow));
  }

  return arrayResult;
}

/**
 * MovingAverageConvergenceDivergence
 * 
 * return an array with MACD
 * 
 * @param {array} arrayInput input array
 * @param {number} arrayColumn index for the value column (usually close price) in the input array
 * @param {number} lookBackFast [OPTIONAL] 12 by default, look back peroid for fast EMA
 * @param {number} lookBackSlow [OPTIONAL] 26 by default, look back period for slow EMA and it should be greater than lookBackFast
 * @param {number} lookBackSignal [OPTIONAL] 9 by default, look back period for signal
 * @return an array with MACD, [fastEMA, slowEMA, MACD, Signal, Histogram]
 * @customfunction
 */
function MACD(arrayInput, arrayColumn, lookBackFast = 12, lookBackSlow = 26, lookBackSignal = 9) {
  // check if input arguments
  if (arguments.length < 2) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBackSlow + lookBackSignal) throw new Error("arrayInput length should be greater");
  if (arrayColumn > arrayInput[0].length - 1) throw new Error("arrayColumn is set out of range");
  if (lookBackFast >= lookBackSlow) throw new Error("lookBackfast should be smaller than lookBackSlow");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;
  var smoothingFast = 2 / (lookBackFast + 1), smoothingSlow = 2 / (lookBackSlow + 1), smoothingSignal = 2 / (lookBackSignal + 1);
  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("fastEMA", "slowEMA", "MACD", "Signal", "Histogram");

  i++;

  //
  // Stage 2: Find the first index to start Fast EMA at
  //
  var arraySkip = 0;
  for (; i < arraySkip + lookBackFast; i++) {
    // no MACD calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][arrayColumn]) arraySkip++;
  }

  var lbStartFast = arraySkip + 1;
  // var indexFastAt = lbStartFast + lookBackFast - 1;

  var lbStartSlow = arraySkip + 1;
  var indextSlowAt = lbStartSlow + lookBackSlow - 1;

  var lbStartSignal = indextSlowAt;
  var indexSignalAt = lbStartSignal + lookBackSignal - 1;


  // the first average calculation with SMA
  arrayResult[i] = arrayInput[i];

  arrayResult[i].push(__HAvg(arrayResult, arrayColumn, lbStartFast, lookBackFast));
  i++

  for (; i < indextSlowAt; i++) {
    arrayResult[i] = arrayInput[i];

    arrayResult[i][rIndex + 1] = smoothingFast * arrayResult[i][arrayColumn] + (1 - smoothingFast) * arrayResult[i - 1][rIndex + 1];
  }

  //
  // Stage 3: Find the first index to start Slow EMA and MACD at
  //
  arrayResult[i] = arrayInput[i];

  var emaFast = smoothingFast * arrayResult[i][arrayColumn] + (1 - smoothingFast) * arrayResult[i - 1][rIndex + 1];
  var emaSlow = __HAvg(arrayResult, arrayColumn, lbStartSlow, lookBackSlow);
  var macd = emaFast - emaSlow;
  arrayResult[i].push(emaFast, emaSlow, macd);
  i++

  // the rest average calcuation with WMA
  for (; i < indexSignalAt; i++) {
    arrayResult[i] = arrayInput[i];

    emaFast = smoothingFast * arrayResult[i][arrayColumn] + (1 - smoothingFast) * arrayResult[i - 1][rIndex + 1];
    emaSlow = smoothingSlow * arrayResult[i][arrayColumn] + (1 - smoothingSlow) * arrayResult[i - 1][rIndex + 2];
    macd = emaFast - emaSlow;
    arrayResult[i].push(emaFast, emaSlow, macd);
  }

  //
  // Stage 4: Find the first index to start Signal at
  //
  arrayResult[i] = arrayInput[i];

  emaFast = smoothingFast * arrayResult[i][arrayColumn] + (1 - smoothingFast) * arrayResult[i - 1][rIndex + 1];
  emaSlow = smoothingSlow * arrayResult[i][arrayColumn] + (1 - smoothingSlow) * arrayResult[i - 1][rIndex + 2];
  macd = emaFast - emaSlow;

  // the first SMA calculation for signal use macd at the current index so update array first
  arrayResult[i].push(emaFast, emaSlow, macd);

  var signal = __HAvg(arrayResult, rIndex + 3, lbStartSignal, lookBackSignal);

  arrayResult[i].push(signal, macd - signal);
  i++

  // the rest calcuation
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    emaFast = smoothingFast * arrayResult[i][arrayColumn] + (1 - smoothingFast) * arrayResult[i - 1][rIndex + 1];
    emaSlow = smoothingSlow * arrayResult[i][arrayColumn] + (1 - smoothingSlow) * arrayResult[i - 1][rIndex + 2];
    macd = emaFast - emaSlow;
    signal = smoothingSignal * macd + (1 - smoothingSignal) * arrayResult[i - 1][rIndex + 4];

    arrayResult[i].push(emaFast, emaSlow, macd, signal, macd - signal);
  }

  return arrayResult;
}

/**
 * AverageDirectionIndex
 * 
 * return an array with ADX
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} lookBackTR [OPTIONAL] 14 by default, look back period for TR, DM+, DM-
 * @param {number} lookBackADX [OPTIONAL] 14 by default, look back period for ADX calculation
 * @return an array with ADX, [DI+, DI-, ADX]
 * @customfunction
 */
function ADX(arrayInput, lookBackTR = 14, lookBackADX = 14) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBackTR + lookBackADX + 1) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  var smoothingTR = 1 / lookBackTR, smoothingADX = 1 / lookBackADX;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("DI+", "DI-", "ADX");
  i++;

  var arrayTrueRange = [], arrayDirMovePlus = [], arrayDirMoveMinus = [];
  //
  // Stage 2: Find the first index to start TR at
  //
  var arraySkip = 0;
  for (; i < lookBackTR + arraySkip + 1; i++) {
    // copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][CLOSE]) { arraySkip++; }
    else {
      // calcualte TR1, DM1+, DM1- and store into array for MA calculation later
      arrayTrueRange.push(Math.max(arrayInput[i + 1][HIGH] - arrayInput[i + 1][LOW], arrayInput[i + 1][HIGH] - arrayInput[i][CLOSE], arrayInput[i][CLOSE] - arrayInput[i + 1][LOW]));
      arrayDirMovePlus.push(arrayInput[i + 1][HIGH] - arrayInput[i][HIGH] > arrayInput[i][LOW] - arrayInput[i + 1][LOW] ? arrayInput[i + 1][HIGH] - arrayInput[i][HIGH] : 0);
      arrayDirMoveMinus.push(arrayInput[i + 1][HIGH] - arrayInput[i][HIGH] < arrayInput[i][LOW] - arrayInput[i + 1][LOW] ? arrayInput[i][LOW] - arrayInput[i + 1][LOW] : 0);
    }
  }

  //
  // Stage 3: Calculate MA of TR14, DM14+, DM14-, DI+, DI-, DX
  //

  //
  // First MA is SMA for TR14, DM14+, DM14-
  //
  arrayResult[i] = arrayInput[i];

  // SMA
  var turRangeMA = __VAvg(arrayTrueRange);
  var dirMovePlusMA = __VAvg(arrayDirMovePlus);
  var dirMoveMinusMA = __VAvg(arrayDirMoveMinus);

  var dirIndexPlus = dirMovePlusMA / turRangeMA;
  var dirIndexMinus = dirMoveMinusMA / turRangeMA;

  var dirIndex = Math.abs(dirIndexPlus - dirIndexMinus) / (dirIndexPlus + dirIndexMinus);

  var arrayDirIndex = [];

  arrayDirIndex.push(dirIndex);

  arrayResult[i].push(dirIndexPlus, dirIndexMinus);

  i++;

  //
  // Wilder's MA for TR14, DM14+, DM14- for the rest
  //
  var trueRange = 0, dirMovePlus = 0, dirMoveMinus = 0;
  for (; i < lookBackTR + lookBackADX; i++) {
    arrayResult[i] = arrayInput[i];

    trueRange = Math.max(arrayResult[i][HIGH] - arrayResult[i][LOW], arrayResult[i][HIGH] - arrayResult[i - 1][CLOSE], arrayResult[i - 1][CLOSE] - arrayResult[i][LOW]);
    dirMovePlus = arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] > arrayResult[i - 1][LOW] - arrayResult[i][LOW] ? arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] : 0;
    dirMoveMinus = arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] < arrayResult[i - 1][LOW] - arrayResult[i][LOW] ? arrayResult[i - 1][LOW] - arrayResult[i][LOW] : 0;

    // Wilder's MA
    turRangeMA = smoothingTR * trueRange + (1 - smoothingTR) * turRangeMA;
    dirMovePlusMA = smoothingTR * dirMovePlus + (1 - smoothingTR) * dirMovePlusMA;
    dirMoveMinusMA = smoothingTR * dirMoveMinus + (1 - smoothingTR) * dirMoveMinusMA;

    dirIndexPlus = dirMovePlusMA / turRangeMA;
    dirIndexMinus = dirMoveMinusMA / turRangeMA;

    dirIndex = Math.abs(dirIndexPlus - dirIndexMinus) / (dirIndexPlus + dirIndexMinus);
    arrayDirIndex.push(dirIndex);

    arrayResult[i].push(dirIndexPlus, dirIndexMinus);
  }

  //
  // Stage 4 calculate ADX, MA of DX
  //

  //
  // First MA is SMA for ADX
  //
  arrayResult[i] = arrayInput[i];

  trueRange = Math.max(arrayResult[i][HIGH] - arrayResult[i][LOW], arrayResult[i][HIGH] - arrayResult[i - 1][CLOSE], arrayResult[i - 1][CLOSE] - arrayResult[i][LOW]);
  dirMovePlus = arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] > arrayResult[i - 1][LOW] - arrayResult[i][LOW] ? arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] : 0;
  dirMoveMinus = arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] < arrayResult[i - 1][LOW] - arrayResult[i][LOW] ? arrayResult[i - 1][LOW] - arrayResult[i][LOW] : 0;

  turRangeMA = smoothingTR * trueRange + (1 - smoothingTR) * turRangeMA;
  dirMovePlusMA = smoothingTR * dirMovePlus + (1 - smoothingTR) * dirMovePlusMA;
  dirMoveMinusMA = smoothingTR * dirMoveMinus + (1 - smoothingTR) * dirMoveMinusMA;

  dirIndexPlus = dirMovePlusMA / turRangeMA;
  dirIndexMinus = dirMoveMinusMA / turRangeMA;

  dirIndex = Math.abs(dirIndexPlus - dirIndexMinus) / (dirIndexPlus + dirIndexMinus); 
  arrayDirIndex.push(dirIndex);

  // SMA
  var averageDirIndex = __VAvg(arrayDirIndex);

  arrayResult[i].push(dirIndexPlus, dirIndexMinus, averageDirIndex);

  i++

  //
  // Wilder's MA for ADX for the rest
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    trueRange = Math.max(arrayResult[i][HIGH] - arrayResult[i][LOW], arrayResult[i][HIGH] - arrayResult[i - 1][CLOSE], arrayResult[i - 1][CLOSE] - arrayResult[i][LOW]);
    dirMovePlus = arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] > arrayResult[i - 1][LOW] - arrayResult[i][LOW] ? arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] : 0;
    dirMoveMinus = arrayResult[i][HIGH] - arrayResult[i - 1][HIGH] < arrayResult[i - 1][LOW] - arrayResult[i][LOW] ? arrayResult[i - 1][LOW] - arrayResult[i][LOW] : 0;

    turRangeMA = smoothingTR * trueRange + (1 - smoothingTR) * turRangeMA;
    dirMovePlusMA = smoothingTR * dirMovePlus + (1 - smoothingTR) * dirMovePlusMA;
    dirMoveMinusMA = smoothingTR * dirMoveMinus + (1 - smoothingTR) * dirMoveMinusMA;

    dirIndexPlus = dirMovePlusMA / turRangeMA;
    dirIndexMinus = dirMoveMinusMA / turRangeMA;

    dirIndex = Math.abs(dirIndexPlus - dirIndexMinus) / (dirIndexPlus + dirIndexMinus);

    // Wilder's MA
    averageDirIndex = smoothingADX * dirIndex + (1 - smoothingADX) * averageDirIndex;

    arrayResult[i].push(dirIndexPlus, dirIndexMinus, averageDirIndex);
  }

  return arrayResult;
}

/**
 * ActiveTrueRange
 * 
 * return an array with ATR
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} lookBack [OPTIONAL] 14 by default, look back period for ATR calculation
 * @param {number} averageMethod 0 for SMA, 1 for EMA, 2 for Wilder's
 * @return an array with ATR, [TR, ATR]
 * @customfunction
 */
function ATR(arrayInput, lookBack = 14, averageMethod = 0) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 2) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("TR", "ATR");
  i++;

  //
  // Stage 2: Find the first index to start ATR at
  //
  var arraySkip = 0, trueRange = 0;
  for (; i < lookBack + arraySkip + 1; i++) {
    // no ATR calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    if (trueRange) arrayResult[i][rIndex + 1] = trueRange;

    // skip empty data index and calculate True Range
    if (!arrayInput[i][CLOSE]) { arraySkip++; }
    else {
      trueRange = Math.max(arrayInput[i + 1][HIGH] - arrayInput[i + 1][LOW], arrayInput[i + 1][HIGH] - arrayInput[i][CLOSE], arrayInput[i][CLOSE] - arrayInput[i + 1][LOW]);
    }
  }

  //
  // Stage 3: Calculate Active True Range
  //

  // the first average calculation with SMA
  arrayResult[i] = arrayInput[i];

  trueRange = Math.max(arrayResult[i][HIGH] - arrayResult[i][LOW], arrayResult[i][HIGH] - arrayResult[i - 1][CLOSE], arrayResult[i - 1][CLOSE] - arrayResult[i][LOW]);
  arrayResult[i][rIndex + 1] = trueRange;

  var lbStart = i - lookBack + 1;
  arrayResult[i][rIndex + 2] = __HAvg(arrayResult, rIndex + 1, lbStart, lookBack);

  i++;

  // the rest calculation according to option { ATRSMA = 0, ATREMA = 1, ATRWilders = 2 }
  var smoothing = (averageMethod == 1) ? 2 / (1 + lookBack) : 1 / lookBack;
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    trueRange = Math.max(arrayResult[i][HIGH] - arrayResult[i][LOW], arrayResult[i][HIGH] - arrayResult[i - 1][CLOSE], arrayResult[i - 1][CLOSE] - arrayResult[i][LOW]);
    arrayResult[i][rIndex + 1] = trueRange;

    switch (averageMethod) {
      case 1:
      case 2:
        arrayResult[i][rIndex + 2] = trueRange * smoothing + arrayResult[i - 1][rIndex + 2] * (1 - smoothing);
        break;

      default:
        var lbStart = i - lookBack + 1;
        arrayResult[i][rIndex + 2] = __HAvg(arrayResult, rIndex + 1, lbStart, lookBack);
    }
  }

  return arrayResult;
}

/**
 * RelativeStrengthIndex
 * 
 * return an array with RSI
 * 
 * @param {array} arrayInput input array
 * @param {number} arrayColumn index for the value column (usually close price) in the input array
 * @param {number} lookBack [OPTIONAL] 14 by default, look back period for RSI calculation
 * @return an array with RSI, [avgGain, avgLoss, RSI]
 * @customfunction
 */
function RSI(arrayInput, arrayColumn, lookBack = 14) {
  // check if input arguments
  if (arguments.length < 2) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 2) throw new Error("arrayInput length should be greater");
  if (arrayColumn > arrayInput[0].length - 1) throw new Error("arrayColumn is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("avgGain", "avgLoss", "RSI");
  i++;

  //
  // Stage 2: Find the first index to start RSI at
  //
  var arraySkip = 0;
  for (; i < arraySkip + lookBack + 1; i++) {
    // no RSI calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][arrayColumn]) arraySkip++;
  }

  //
  // Stage 3: Calculate Relative Strength Index
  //

  // the first average calculation with SMA
  arrayResult[i] = arrayInput[i];

  var lbStart = i - lookBack;
  var relStrIndex = __RSISMA(arrayResult, arrayColumn, lbStart, lookBack);
  arrayResult[i].push(relStrIndex[0], relStrIndex[1], relStrIndex[2]);
  i++

  // the rest average calcuation with WMA
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    relStrIndex = __RSIWMA(arrayResult, arrayColumn, i, lookBack, rIndex);

    arrayResult[i].push(relStrIndex[0], relStrIndex[1], relStrIndex[2]);
  }

  return arrayResult;
}

function __RSISMA(arrayInput, arrayColumn, lbStart, lookBack) {
  var sumGain = 0, sumLoss = 0;
  for (var i = lbStart + 1; i < lbStart + lookBack + 1; i++) {
    var change = arrayInput[i][arrayColumn] - arrayInput[i - 1][arrayColumn];

    sumGain += (change > 0 ? change : 0);
    sumLoss += (change < 0 ? -1 * change : 0);
  }
  var avgGain = sumGain / lookBack;
  var avgLoss = sumLoss / lookBack;
  var relStr = (avgLoss == 0 ? 100 : avgGain / avgLoss);

  return [avgGain, avgLoss, (100 - (100 / (1 + relStr)))];
}

function __RSIWMA(arrayInput, arrayColumn, lbStart, lookBack, rIndex) {
  var change = arrayInput[lbStart][arrayColumn] - arrayInput[lbStart - 1][arrayColumn];

  var prevGain = arrayInput[lbStart - 1][rIndex + 1];
  var prevLoss = arrayInput[lbStart - 1][rIndex + 2];

  var avgGain = (change > 0 ? ((prevGain * (lookBack - 1) + change) / lookBack) : (prevGain * (lookBack - 1)) / lookBack);
  var avgLoss = (change < 0 ? ((prevLoss * (lookBack - 1) - change) / lookBack) : (prevLoss * (lookBack - 1)) / lookBack);

  var relStr = (avgLoss == 0 ? 100 : avgGain / avgLoss);

  return [avgGain, avgLoss, (100 - (100 / (1 + relStr)))];
}

/**
 * WilliamVixFix
 * 
 * return an array with WVF and bolinger band for WVF
 *
 * WVF = (Highest(close,22) - Low) / Highest(close,22) * 100
 * WVFVM(volume modified) = (Highest(close, 22) - Low) * Volume / (Highest(close, 22) * Average(volume, 22)) * 100
 * 
 * @param {array} arrayInput input array, firt 6 column of the input array should be DOHLCV format (Date, Open, High, Low, Close, Volume)
 * @param {number} wvfLength [OPTIONAL] 22 by default, look back period for WVF
 * @param {number} bandLength [OPTIONAL] 20 by default, look back period for bolinger band for WVF
 * @param {number} stDitance [OPTIONAL] 2.0 by default, standard deviation for high band and low band width calculation for bolinger band  
 * @param {bool} volumeModified [OPTIONAL] false by default, WVF with volume modified
 * @return an array with WVF, [WVF, midLine, lowBand, highBand]
 * @customfunction
 */
function WVF(arrayInput, wvfLength = 22, bandLength = 20, stDitance = 2.0, volumeModified = false) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < wvfLength + 1) throw new Error("arrayInput length should be greater");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("WVF");
  i++;

  //
  // Stage 2: Find the first index to start WVF at
  //
  var arraySkip = 0;
  for (; i < wvfLength + arraySkip; i++) {
    // no WVF calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][CLOSE]) arraySkip++;
  }

  //
  // Stage 3: Calculate William Vix Fix
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    var lbStart = i - wvfLength + 1;

    var highest = __HMax(arrayResult, HIGH, lbStart, wvfLength);
    var volumeAvg = volumeModified ? __HAvg(arrayResult, VOLUME, lbStart, wvfLength) : 1;
    var volumeLow = volumeModified ? arrayResult[i][VOLUME] : 1;

    if (volumeLow == 0) volumeLow = volumeAvg = 1;

    var wvf = (highest - arrayResult[i][LOW]) * volumeLow / (highest * volumeAvg) * 100;

    arrayResult[i].push(wvf);
  }

  return (wvfLength + bandLength <= arrayInput.length ? BolingerBand(arrayResult, rIndex + 1, bandLength, stDitance) : arrayResult);
}

/**
 * Bolinger Band
 * 
 * return an array with bolinger band
 * 
 * @param {array} arrayInput input array
 * @param {number} arrayColumn index for the value column (usually close price) in the input array
 * @param {number} lookBack look back period for bolinger band calculation
 * @param {number} stDitance standard deviation to calculate high band and low band
 * @return an array with bolinger band, [midLine, lowBand, highBand]
 * @customfunction
 */
function BolingerBand(arrayInput, arrayColumn, lookBack, stDitance) {
  // check if input arguments
  if (arguments.length < 4) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 1) throw new Error("arrayInput length should be greater");
  if (arrayColumn > arrayInput[0].length - 1) throw new Error("arrayColumn is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("midLine", "lowBand", "highBand");
  i++;

  //
  // Stage 2: Find the first index to start Bolinger Band at
  //
  var arraySkip = 0;
  for (; i < lookBack + arraySkip; i++) {
    // no bolinger band calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][arrayColumn]) arraySkip++;
  }

  //
  // Stage 3: Calculate Bolinger Band
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    var lbStart = i - lookBack + 1;

    var sDev = stDitance * __HStdev(arrayResult, arrayColumn, lbStart, lookBack);
    var avg = __HAvg(arrayResult, arrayColumn, lbStart, lookBack);

    arrayResult[i].push(avg, avg - sDev, avg + sDev);
  }

  return arrayResult;
}

/**
 * RateOfChange
 * 
 * return an array with ROC
 * 
 * @param {array} arrayInput input array
 * @param {number} arrayColumn index for the value column (usually close price) in the input array
 * @param {number} lookBack [OPTIONAL] 1 by default, look back period
 * @param {bool} logReturn [OPTIONAL] false by default, true for LOG return calculation, LN(currentPrice/prevPrice),
 * @return an array of with ROC, [ROC]
 * @customfunction
 */
function ROC(arrayInput, arrayColumn, lookBack = 1, logReturn = false) {
  // check if input arguments
  if (arguments.length < 2) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < 3) throw new Error("arrayInput length should be greater");
  if (arrayColumn > arrayInput[0].length - 1) throw new Error("arrayColumn is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  // add result column headers
  arrayResult[0].push("ROC");

  i++;

  //
  // Stage 2: Find the first index to start ROC at
  //
  var arraySkip = 0;
  for (; i < arraySkip + lookBack + 1; i++) {
    // no calculation, just copy input array over
    arrayResult[i] = arrayInput[i];

    // skip empty data index
    if (!arrayInput[i][arrayColumn]) arraySkip++;
  }

  //
  // Stage 3: Calculate Return of Change
  //
  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    arrayResult[i].push(logReturn ? Math.log(arrayResult[i][arrayColumn] / arrayResult[i - lookBack][arrayColumn]) : (arrayResult[i][arrayColumn] / arrayResult[i - lookBack][arrayColumn]) - 1);
  }

  return arrayResult;
}

/**
 * FlipArray
 * Flip array upside down
 * 
 * @param {array} arrayInput input array
 * @param {bool} headerRow [OPTIONAL] true by default, false if no headers low exist
 * @return an array upside down
 * @customfunction
 */
function FlipArray(arrayInput, headerRow = true) {
  // check if input arguments
  if (arguments.length < 1) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");

  var i = 0, arrayResult = [];

  if (headerRow) {
    // copy header row
    arrayResult[0] = arrayInput[0];
    i++;
  }

  for (; i < arrayInput.length; i++) {
    var inverted = arrayInput.length - i;
    arrayResult.push(arrayInput[inverted]);
  }
  return arrayResult;
}

/**
 * HistoricalVolitality
 * 
 * return an array with historical volitality
 * 
 * @param {array} arrayInput input array
 * @param {number} arrayColumn index for the value column (usually close price) in the input array
 * @param {number} lookBack 22 by default, look back period for HV calculation
 * @param {bool} logReturn [OPTIONAL] false by default, true for LOG return calculation, LN(currentPrice/prevPrice),
 * @return an array with HV, [HV]
 * @customfunction
 */
function HV(arrayInput, arrayColumn, lookBack = 22, logReturn = false, population = false) {
  // check if input arguments
  if (arguments.length < 2) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 2) throw new Error("arrayInput length should be greater");
  if (arrayColumn > arrayInput[0].length - 1) throw new Error("arrayColumn is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  // get rate of return first
  arrayResult = ROC(arrayInput, arrayColumn, 1, logReturn);

  // add header
  arrayResult[0].push("HV");

  var arraySkip = 0;
  // skip empty data index
  for (i = 1; i < arraySkip + 2; i++) {
    if (!arrayResult[i][rIndex + 1]) arraySkip++;
  }

  var lbStart = arraySkip + 1;
  var indexAt = lbStart + lookBack - 1;

  for (; i < indexAt; i++);

  for (; i < arrayResult.length; i++) {
    arrayResult[i].push(__HStdev(arrayResult, rIndex + 1, i - lookBack + 1, lookBack));
  }

  return arrayResult;
}

/**
 * SimpleMovingAverage
 * 
 * return an array with SMA
 * 
 * @param {array} arrayInput input array
 * @param {number} arrayColumn index for the value column (usually close price) in the input array
 * @param {number} lookBack look back period for SMA calculation
 * @return an array with SMA, [SMA]
 * @customfunction
 */
function SMA(arrayInput, arrayColumn, lookBack) {
  // check if input arguments
  if (arguments.length < 3) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 1) throw new Error("arrayInput length should be greater");
  if (arrayColumn > arrayInput[0].length - 1) throw new Error("arrayColumn is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  arrayResult[0].push("SMA" + lookBack);
  i++;

  var arraySkip = 0;
  // skip empty data index
  for (i = 1; i < arraySkip + 2; i++) {
    arrayResult[i] = arrayInput[i];

    if (!arrayResult[i][arrayColumn]) arraySkip++;
  }

  var lbStart = arraySkip + 1;
  var indexAt = lbStart + lookBack - 1;

  for (; i < indexAt; i++) {
    arrayResult[i] = arrayInput[i];
  }

  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];
    arrayResult[i].push(__HAvg(arrayResult, arrayColumn, i - lookBack + 1, lookBack));
  }

  return arrayResult;
}

/**
 * ExponentMovingAverage
 * 
 * return an array with EMA
 * 
 * @param {array} arrayInput input array
 * @param {number} arrayColumn index for the value column (usually close price) in the input array
 * @param {number} lookBack look back period for EMA calculation
 * @return an array with EMA, [EMA]
 * @customfunction
 */
function EMA(arrayInput, arrayColumn, lookBack) {
  // check if input arguments
  if (arguments.length < 3) throw new Error("more arguments are required");
  if (arrayInput.map == null) throw new Error("arrayInput is not an array");
  if (arrayInput.length < lookBack + 1) throw new Error("arrayInput length should be greater");
  if (arrayColumn > arrayInput[0].length - 1) throw new Error("arrayColumn is set out of range");

  var i = 0, arrayResult = [], rIndex = arrayInput[0].length - 1;
  var smoothing = 2 / (1 + lookBack);

  //
  // Stage 1: Header copy
  // 
  // copy column headers
  arrayResult[0] = arrayInput[0];

  arrayResult[0].push("EMA" + lookBack);
  i++;

  var arraySkip = 0;
  // skip empty data index
  for (i = 1; i < arraySkip + 2; i++) {
    arrayResult[i] = arrayInput[i];

    if (!arrayResult[i][arrayColumn]) arraySkip++;
  }

  var lbStart = arraySkip + 1;
  var indexAt = lbStart + lookBack - 1;

  for (; i < indexAt; i++) {
    arrayResult[i] = arrayInput[i];
  }

  // first average using SMA
  arrayResult[i] = arrayInput[i];
  arrayResult[i].push(__HAvg(arrayResult, arrayColumn, i - lookBack + 1, lookBack));
  i++;

  for (; i < arrayInput.length; i++) {
    arrayResult[i] = arrayInput[i];

    arrayResult[i].push((smoothing * arrayResult[i][arrayColumn]) + (1 - smoothing) * arrayResult[i - 1][rIndex + 1]);
  }

  return arrayResult;
}

/*
 * 
 * Utilities
 * 
**/

// OHLCV index
const DATE = 0, OPEN = 1, HIGH = 2, LOW = 3, CLOSE = 4, VOLUME = 5;


// convert array-value [] into object-key-value {}
function OptionObject(arrayOptions) {
  var objOptions = {};
  for (var i in arrayOptions) {
    if (objOptions[arrayOptions[i][0]]) throw new Error("duplicated key is used for option object");
    objOptions[arrayOptions[i][0]] = arrayOptions[i][1];
  }
  return objOptions;
}

// horizontal covariance
function __HCovar(arrayFirst, colFirst, arraySecond, colSecond, start, length, population = false) {
  var arrayTempFirst = [], arrayTempSecond = [];
  for (var i = start; i < start + length; i++) {
    arrayTempFirst[i - start] = arrayFirst[i][colFirst];
    arrayTempSecond[i - start] = arraySecond[i][colSecond];
  }

  return __VCovar(arrayTempFirst, arrayTempSecond, population);
}

// Calculate covariance, sample or population
function __VCovar(arrayFirst, arraySecond, population = false) {
  var sum = 0;
  var meanFirst = __VAvg(arrayFirst, true);
  var meanSecond = __VAvg(arraySecond, true);

  for (var i in arrayFirst) {
    sum += (arrayFirst[i] - meanFirst) * (arraySecond[i] - meanSecond);
  }

  return sum / (population ? arrayFirst.length : arrayFirst.length - 1);
}

function __HStdev(arrayInput, column, start, length, population = false) {
  var arrayTemp = [];
  for (var i = start; i < start + length; i++) {
    arrayTemp[i - start] = arrayInput[i][column];
  }

  return __VStdev(arrayTemp, population);
}

// Calculate standard deviation, sample or population
function __VStdev(arrayInput, population = false) {
  return Math.sqrt(__VVar(arrayInput, population));
}

function __HVar(arrayInput, column, start, length, population = false) {
  var arrayTemp = [];
  for (var i = start; i < start + length; i++) {
    arrayTemp[i - start] = arrayInput[i][column];
  }

  return __VVar(arrayTemp, population);
}

// Calculate variance, sample or population
function __VVar(arrayInput, population = false) {
  var mean = __VAvg(arrayInput, true);

  return __VAvg(arrayInput.map(function (sum) { return Math.pow(sum - mean, 2); }), population);
}

function __HAvg(arrayInput, column, start, length, population = true) {
  return __HSum(arrayInput, column, start, length) / (population ? length : length - 1);
}

function __VAvg(arrayInput, population = true) {
  return __VSum(arrayInput) / (population ? arrayInput.length : arrayInput.length - 1);
}

function __HSum(arrayInput, column, start, length) {
  var sum = 0;

  for (var i = start; i < start + length; i++) {
    sum += 1 * arrayInput[i][column];
  }

  return sum;
}

function __VSum(arrayInput) {
  var sum = 0;

  for (var i in arrayInput) sum += 1 * arrayInput[i];

  return sum;
}

function __HMax(arrayInput, column, start, length) {
  var max = 1 * arrayInput[start][column];

  for (var i = start + 1; i < start + length; i++) {
    if (arrayInput[i][column] > max) max = 1 * arrayInput[i][column];
  }

  return max;
}

function __HMin(arrayInput, column, start, length) {
  var min = 1 * arrayInput[start][column];

  for (var i = start + 1; i < start + length; i++) {
    if (arrayInput[i][column] < min) min = 1 * arrayInput[i][column];
  }

  return min;
}
