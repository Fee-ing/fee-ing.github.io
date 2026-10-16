/* Ledger — function catalogue for Insert Function, Function Arguments and the formula ScreenTips.
 * Format per line: NAME|arg1;arg2;[optional];more...|Description
 *
 * Derived from Microsoft VBA documentation (github.com/MicrosoftDocs/VBA-Docs, the Excel WorksheetFunction pages;
 * Creative Commons Attribution 4.0, (c) Microsoft Corporation; common/licenses/microsoft-vba-docs-LICENSE.txt):
 * the descriptions and the argument help follow its text, in places shortened or adapted to Ledger. */
(function (root) {
  'use strict';
  const L = root.L;
  const FI = (L.fninfo = {});
  const DATA = {
    Financial: `
ACCRINT|issue;first_interest;settlement;rate;par;frequency;[basis];[calc_method]|返回定期付息证券的应计利息。
ACCRINTM|issue;settlement;rate;par;[basis]|返回到期付息证券的应计利息。
AMORDEGRC|cost;date_purchased;first_period;salvage;period;rate;[basis]|使用折旧系数，返回资产在各会计期间按比例分摊的线性折旧额。
AMORLINC|cost;date_purchased;first_period;salvage;period;rate;[basis]|返回资产在各会计期间按比例分摊的线性折旧额。
COUPDAYBS|settlement;maturity;frequency;[basis]|返回付息期起始日到结算日之间的天数。
COUPDAYS|settlement;maturity;frequency;[basis]|返回包含结算日的付息期的天数。
COUPDAYSNC|settlement;maturity;frequency;[basis]|返回结算日到下一个付息日之间的天数。
COUPNCD|settlement;maturity;frequency;[basis]|返回结算日之后的下一个付息日。
COUPNUM|settlement;maturity;frequency;[basis]|返回结算日与到期日之间的付息次数。
COUPPCD|settlement;maturity;frequency;[basis]|返回结算日之前的上一个付息日。
CUMIPMT|rate;nper;pv;start_period;end_period;type|返回两个期间之间累计支付的利息。
CUMPRINC|rate;nper;pv;start_period;end_period;type|返回两个期间之间累计偿还的贷款本金。
DB|cost;salvage;life;period;[month]|使用固定余额递减法，返回资产在指定期间的折旧额。
DDB|cost;salvage;life;period;[factor]|使用双倍余额递减法或其他指定方法，返回资产在指定期间的折旧额。
DISC|settlement;maturity;pr;redemption;[basis]|返回证券的贴现率。
DOLLARDE|fractional_dollar;fraction|将以分数表示的美元价格转换为以小数表示的美元价格。
DOLLARFR|decimal_dollar;fraction|将以小数表示的美元价格转换为以分数表示的美元价格。
DURATION|settlement;maturity;coupon;yld;frequency;[basis]|返回定期付息证券的年久期。
EFFECT|nominal_rate;npery|返回实际年利率。
FV|rate;nper;pmt;[pv];[type]|根据定期定额付款和固定利率，返回投资的终值。
FVSCHEDULE|principal;schedule|对初始本金应用一系列复利利率，返回其终值。
INTRATE|settlement;maturity;investment;redemption;[basis]|返回全额投资证券的利率。
IPMT|rate;per;nper;pv;[fv];[type]|根据定期定额付款和固定利率，返回投资在指定期间的利息付款额。
IRR|values;[guess]|返回一系列现金流的内部收益率。
ISPMT|rate;per;nper;pv|返回投资在指定期间支付的利息。
MDURATION|settlement;maturity;coupon;yld;frequency;[basis]|返回假设面值为 $100 的证券的修正麦考利久期。
MIRR|values;finance_rate;reinvest_rate|返回一系列定期现金流的内部收益率，同时考虑投资成本和现金再投资的利息。
NOMINAL|effect_rate;npery|返回名义年利率。
NPER|rate;pmt;pv;[fv];[type]|根据定期定额付款和固定利率，返回投资的期数。
NPV|rate;value1;[value2];...|根据贴现率及一系列未来付款（负值）和收入（正值），返回投资的净现值。
ODDFPRICE|settlement;maturity;issue;first_coupon;rate;yld;redemption;frequency;[basis]|返回每 $100 面值、首期不规则的证券的价格。
ODDFYIELD|settlement;maturity;issue;first_coupon;rate;pr;redemption;frequency;[basis]|返回首期不规则的证券的收益率。
ODDLPRICE|settlement;maturity;last_interest;rate;yld;redemption;frequency;[basis]|返回每 $100 面值、末期不规则的证券的价格。
ODDLYIELD|settlement;maturity;last_interest;rate;pr;redemption;frequency;[basis]|返回末期不规则的证券的收益率。
PDURATION|rate;pv;fv|返回投资达到指定价值所需的期数。
PMT|rate;nper;pv;[fv];[type]|根据固定还款额和固定利率计算贷款的每期付款额。
PPMT|rate;per;nper;pv;[fv];[type]|根据定期定额付款和固定利率，返回指定投资的本金付款额。
PRICE|settlement;maturity;rate;yld;redemption;frequency;[basis]|返回每 $100 面值定期付息证券的价格。
PRICEDISC|settlement;maturity;discount;redemption;[basis]|返回每 $100 面值贴现证券的价格。
PRICEMAT|settlement;maturity;issue;rate;yld;[basis]|返回每 $100 面值到期付息证券的价格。
PV|rate;nper;pmt;[fv];[type]|返回投资的现值，即未来一系列付款在当前的总价值。
RATE|nper;pmt;pv;[fv];[type];[guess]|返回贷款或投资的每期利率。
RECEIVED|settlement;maturity;investment;discount;[basis]|返回全额投资证券在到期时收到的金额。
RRI|nper;pv;fv|返回投资增长的等效利率。
SLN|cost;salvage;life|返回资产在一个期间内按直线法计算的折旧额。
SYD|cost;salvage;life;per|返回资产在指定期间按年数总和法计算的折旧额。
TBILLEQ|settlement;maturity;discount|返回国库券的债券等效收益率。
TBILLPRICE|settlement;maturity;discount|返回每 $100 面值国库券的价格。
TBILLYIELD|settlement;maturity;pr|返回国库券的收益率。
VDB|cost;salvage;life;start_period;end_period;[factor];[no_switch]|使用双倍余额递减法或其他指定方法，返回资产在任意指定期间（包括不足一期）的折旧额。
XIRR|values;dates;[guess]|返回不一定定期发生的一系列现金流的内部收益率。
XNPV|rate;values;dates|返回不一定定期发生的一系列现金流的净现值。
YIELD|settlement;maturity;rate;pr;redemption;frequency;[basis]|返回定期付息证券的收益率。
YIELDDISC|settlement;maturity;pr;redemption;[basis]|返回贴现证券（例如国库券）的年收益率。
YIELDMAT|settlement;maturity;issue;rate;pr;[basis]|Returns the annual yield of a security that pays interest at maturity.`,
    'Date & Time': `
DATE|year;month;day|返回表示日期的 Ledger 日期时间序列号。
DATEDIF|start_date;end_date;unit|计算两个日期之间的天数、月数或年数。
DATEVALUE|date_text|将文本形式的日期转换为 Ledger 日期时间代码中的日期序列号。
DAY|serial_number|返回月份中的日，数值为 1 至 31。
DAYS|end_date;start_date|返回两个日期之间的天数。
DAYS360|start_date;end_date;[method]|按一年 360 天（每月 30 天，共十二个月）计算两个日期之间的天数。
EDATE|start_date;months|返回起始日期之前或之后指定月数的日期序列号。
EOMONTH|start_date;months|返回指定月数之前或之后月份的最后一天的日期序列号。
HOUR|serial_number|返回小时，数值为 0（凌晨 12 点）至 23（晚上 11 点）。
ISOWEEKNUM|date|返回指定日期所在年份中的 ISO 周数。
MINUTE|serial_number|返回分钟，数值为 0 至 59。
MONTH|serial_number|返回月份，数值为 1（一月）至 12（十二月）。
NETWORKDAYS|start_date;end_date;[holidays]|返回两个日期之间的完整工作日数。
NETWORKDAYS.INTL|start_date;end_date;[weekend];[holidays]|使用自定义周末参数，返回两个日期之间的完整工作日数。
NOW||返回当前日期和时间，并设置为日期时间格式。
SECOND|serial_number|返回秒数，数值为 0 至 59。
TIME|hour;minute;second|将以数字给出的时、分、秒转换为 Ledger 序列号，并设置为时间格式。
TIMEVALUE|time_text|将文本形式的时间转换为 Ledger 时间序列号，范围为 0（凌晨 12:00:00）至 0.999988426（晚上 11:59:59）。
TODAY||返回当前日期，并设置为日期格式。
WEEKDAY|serial_number;[return_type]|返回 1 至 7 的数字，表示日期是一周中的第几天。
WEEKNUM|serial_number;[return_type]|返回一年中的周数。
WORKDAY|start_date;days;[holidays]|返回指定工作日数之前或之后的日期序列号。
WORKDAY.INTL|start_date;days;[weekend];[holidays]|使用自定义周末参数，返回指定工作日数之前或之后的日期序列号。
YEAR|serial_number|返回日期的年份，数值为 1900 至 9999 的整数。
YEARFRAC|start_date;end_date;[basis]|Returns the year fraction representing the number of whole days between start_date and end_date.`,
    'Math & Trig': `
ABS|number|返回数字的绝对值，即不带符号的数值。
ACOS|number|返回数字的反余弦值，以弧度表示，范围为 0 至 Pi。
ACOSH|number|返回数字的反双曲余弦值。
ACOT|number|返回数字的反余切值，以弧度表示，范围为 0 至 Pi。
ACOTH|number|返回数字的反双曲余切值。
AGGREGATE|function_num;options;ref1;...|返回列表或数据库中的汇总值，可选择忽略隐藏行和错误值。
ARABIC|text|将罗马数字转换为阿拉伯数字。
ASIN|number|返回数字的反正弦值，以弧度表示，范围为 -Pi/2 至 Pi/2。
ASINH|number|返回数字的反双曲正弦值。
ATAN|number|返回数字的反正切值，以弧度表示，范围为 -Pi/2 至 Pi/2。
ATAN2|x_num;y_num|返回指定 x、y 坐标的反正切值，以弧度表示，范围为 -Pi 至 Pi，不包括 -Pi。
ATANH|number|返回数字的反双曲正切值。
BASE|number;radix;[min_length]|将数值转换为指定进制的文本表示。
CEILING|number;significance|将数字向上舍入到最接近的指定基数的倍数。
CEILING.MATH|number;[significance];[mode]|将数字向上舍入到最接近的整数或指定基数的倍数。
COMBIN|number;number_chosen|返回指定数目项目的组合数。
COMBINA|number;number_chosen|返回指定数目项目允许重复的组合数。
COS|number|返回角度的余弦值。
COSH|number|返回数字的双曲余弦值。
COT|number|返回角度的余切值。
COTH|number|返回数字的双曲余切值。
CSC|number|返回角度的余割值。
CSCH|number|返回角度的双曲余割值。
DECIMAL|number;radix|将以指定进制表示的数字文本转换为十进制数。
DEGREES|angle|将弧度转换为角度。
EVEN|number|将正数向上、负数向下舍入到最接近的偶数整数。
EXP|number|返回 e 的指定次幂。
FACT|number|返回数字的阶乘，等于 1*2*3*...*Number。
FACTDOUBLE|number|返回数字的双阶乘。
FLOOR|number;significance|将数字向下舍入到最接近的指定基数的倍数。
FLOOR.MATH|number;[significance];[mode]|将数字向下舍入到最接近的整数或指定基数的倍数。
GCD|number1;[number2];...|返回最大公约数。
INT|number|将数字向下舍入到最接近的整数。
LCM|number1;[number2];...|返回最小公倍数。
LN|number|返回数字的自然对数。
LOG|number;[base]|返回数字以指定数为底的对数。
LOG10|number|返回数字以 10 为底的对数。
MDETERM|array|返回数组矩阵的行列式。
MINVERSE|array|返回数组所存矩阵的逆矩阵。
MMULT|array1;array2|返回两个数组的矩阵乘积；结果的行数与 array1 相同，列数与 array2 相同。
MOD|number;divisor|返回数字除以除数后的余数。
MROUND|number;multiple|返回舍入到指定倍数的数字。
MULTINOMIAL|number1;[number2];...|返回一组数字的多项式系数。
MUNIT|dimension|返回指定维度的单位矩阵。
ODD|number|将正数向上、负数向下舍入到最接近的奇数整数。
PI||返回 Pi 的值 3.14159265358979，精确到 15 位数字。
POWER|number;power|返回数字的指定次幂。
PRODUCT|number1;[number2];...|将参数中的所有数字相乘。
QUOTIENT|numerator;denominator|返回除法商的整数部分。
RADIANS|angle|将角度转换为弧度。
RAND||返回大于等于 0 且小于 1 的均匀分布随机数，重新计算时会更改。
RANDARRAY|[rows];[columns];[min];[max];[integer]|返回随机数数组。
RANDBETWEEN|bottom;top|返回指定两个数字之间的随机数。
ROMAN|number;[form]|将阿拉伯数字转换为文本形式的罗马数字。
ROUND|number;num_digits|将数字舍入到指定的位数。
ROUNDDOWN|number;num_digits|将数字朝零的方向向下舍入。
ROUNDUP|number;num_digits|将数字向远离零的方向向上舍入。
SEC|number|返回角度的正割值。
SECH|number|返回角度的双曲正割值。
SEQUENCE|rows;[columns];[start];[step]|返回数字序列。
SERIESSUM|x;n;m;coefficients|根据公式返回幂级数之和。
SIGN|number|返回数字的符号：正数返回 1，零返回 0，负数返回 -1。
SIN|number|返回角度的正弦值。
SINH|number|返回数字的双曲正弦值。
SQRT|number|返回数字的平方根。
SQRTPI|number|返回 number * Pi 的平方根。
SUBTOTAL|function_num;ref1;...|返回列表或数据库中的分类汇总。
SUM|number1;[number2];...|计算单元格区域中所有数值的总和。
SUMIF|range;criteria;[sum_range]|对符合指定条件的单元格求和。
SUMIFS|sum_range;criteria_range1;criteria1;...|对符合一组指定条件的单元格求和。
SUMPRODUCT|array1;[array2];...|返回对应区域或数组中各元素乘积的总和。
SUMSQ|number1;[number2];...|返回参数的平方和。参数可以是数字、数组、名称或包含数字的单元格引用。
SUMX2MY2|array_x;array_y|返回两个对应区域或数组中各对数值平方差的总和。
SUMX2PY2|array_x;array_y|返回两个对应区域或数组中各对数值的平方和的总和。
SUMXMY2|array_x;array_y|返回两个对应区域或数组中各对数值之差的平方和。
TAN|number|返回角度的正切值。
TANH|number|返回数字的双曲正切值。
TRUNC|number;[num_digits]|Truncates a number to an integer by removing the decimal, or fractional, part of the number.`,
    Statistical: `
AVEDEV|number1;[number2];...|返回数据点与其平均值的绝对偏差的平均值。
AVERAGE|number1;[number2];...|返回参数的平均值（算术平均值）；参数可以是数字，或包含数字的名称、数组及引用。
AVERAGEA|value1;[value2];...|返回参数的平均值（算术平均值），其中的文本和 FALSE 视为 0，TRUE 视为 1。
AVERAGEIF|range;criteria;[average_range]|计算符合给定条件的单元格的平均值（算术平均值）。
AVERAGEIFS|average_range;criteria_range1;criteria1;...|计算符合一组给定条件的单元格的平均值（算术平均值）。
BETA.DIST|x;alpha;beta;cumulative;[A];[B]|返回 Beta 概率分布函数。
BETA.INV|probability;alpha;beta;[A];[B]|返回累积 Beta 概率密度函数的反函数值。
BETADIST|x;alpha;beta;[A];[B]|返回累积 Beta 概率密度函数。
BETAINV|probability;alpha;beta;[A];[B]|返回累积 Beta 概率密度函数的反函数值。
BINOM.DIST|number_s;trials;probability_s;cumulative|返回二项分布的单项概率。
BINOM.DIST.RANGE|trials;probability_s;number_s;[number_s2]|使用二项分布返回试验结果的概率。
BINOM.INV|trials;probability_s;alpha|返回使累积二项分布大于等于指定阈值的最小值。
BINOMDIST|number_s;trials;probability_s;cumulative|返回二项分布的单项概率。
CHIDIST|x;deg_freedom|返回卡方分布的单尾概率。
CHIINV|probability;deg_freedom|返回卡方分布单尾概率的反函数值。
CHISQ.DIST|x;deg_freedom;cumulative|返回卡方分布的左尾概率。
CHISQ.DIST.RT|x;deg_freedom|返回卡方分布的右尾概率。
CHISQ.INV|probability;deg_freedom|返回卡方分布左尾概率的反函数值。
CHISQ.INV.RT|probability;deg_freedom|返回卡方分布右尾概率的反函数值。
CHISQ.TEST|actual_range;expected_range|返回独立性检验结果。
CHITEST|actual_range;expected_range|返回独立性检验结果，即统计量和相应自由度对应的卡方分布值。
CONFIDENCE|alpha;standard_dev;size|返回总体平均值的置信区间。
CONFIDENCE.NORM|alpha;standard_dev;size|使用正态分布返回总体平均值的置信区间。
CONFIDENCE.T|alpha;standard_dev;size|使用 Student t 分布返回总体平均值的置信区间。
CORREL|array1;array2|返回两个数据集的相关系数。
COUNT|value1;[value2];...|统计区域内包含数字的单元格数量。
COUNTA|value1;[value2];...|统计区域内非空单元格的数量。
COUNTBLANK|range|统计指定单元格区域内的空单元格数量。
COUNTIF|range;criteria|统计区域内符合给定条件的单元格数量。
COUNTIFS|criteria_range1;criteria1;...|统计满足一组给定条件的单元格数量。
COVAR|array1;array2|返回协方差，即两个数据集中每对数据点的离差乘积的平均值。
COVARIANCE.P|array1;array2|返回总体协方差，即两个数据集中每对数据点的离差乘积的平均值。
COVARIANCE.S|array1;array2|返回样本协方差，即两个数据集中每对数据点的离差乘积的平均值。
CRITBINOM|trials;probability_s;alpha|返回使累积二项分布大于等于指定阈值的最小值。
DEVSQ|number1;[number2];...|返回各数据点与样本平均值之差的平方和。
EXPON.DIST|x;lambda;cumulative|返回指数分布。
EXPONDIST|x;lambda;cumulative|返回指数分布。
F.DIST|x;deg_freedom1;deg_freedom2;cumulative|返回两个数据集的左尾 F 概率分布（差异程度）。
F.DIST.RT|x;deg_freedom1;deg_freedom2|返回两个数据集的右尾 F 概率分布（差异程度）。
F.INV|probability;deg_freedom1;deg_freedom2|返回左尾 F 概率分布的反函数值。
F.INV.RT|probability;deg_freedom1;deg_freedom2|返回右尾 F 概率分布的反函数值。
F.TEST|array1;array2|返回 F 检验结果，即 Array1 和 Array2 的方差没有显著差异的双尾概率。
FDIST|x;deg_freedom1;deg_freedom2|返回两个数据集的 F 概率分布（差异程度）。
FINV|probability;deg_freedom1;deg_freedom2|返回 F 概率分布的反函数值。
FISHER|x|返回 Fisher 变换值。
FISHERINV|y|返回 Fisher 变换的反函数值。
FORECAST|x;known_y's;known_x's|根据现有值，沿线性趋势计算或预测未来值。
FORECAST.LINEAR|x;known_y's;known_x's|根据现有值，沿线性趋势计算或预测未来值。
FREQUENCY|data_array;bins_array|计算数值在各区间内出现的频率，返回一个垂直数组，其元素数比 Bins_array 多一个。
FTEST|array1;array2|返回 F 检验结果，即 Array1 和 Array2 的方差没有显著差异的双尾概率。
GAMMA|x|返回 Gamma 函数值。
GAMMA.DIST|x;alpha;beta;cumulative|返回 Gamma 分布。
GAMMA.INV|probability;alpha;beta|返回 Gamma 累积分布的反函数值。
GAMMADIST|x;alpha;beta;cumulative|返回 Gamma 分布。
GAMMAINV|probability;alpha;beta|返回 Gamma 累积分布的反函数值。
GAMMALN|x|返回 Gamma 函数的自然对数。
GAMMALN.PRECISE|x|返回 Gamma 函数的自然对数。
GAUSS|x|返回标准正态累积分布值减去 0.5 的结果。
GEOMEAN|number1;[number2];...|返回正数数组或区域的几何平均值。
GROWTH|known_y's;[known_x's];[new_x's];[const]|返回拟合已知数据点的指数增长趋势值。
HARMEAN|number1;[number2];...|返回正数数据集的调和平均值，即各数倒数的算术平均值的倒数。
HYPGEOM.DIST|sample_s;number_sample;population_s;number_pop;cumulative|返回超几何分布。
HYPGEOMDIST|sample_s;number_sample;population_s;number_pop|返回超几何分布。
INTERCEPT|known_y's;known_x's|根据已知的 x 值和 y 值拟合回归直线，计算该直线与 y 轴的交点。
KURT|number1;[number2];...|返回数据集的峰度。
LARGE|array;k|返回数据集中第 k 个最大值，例如第五大的数字。
LINEST|known_y's;[known_x's];[const];[stats]|使用最小二乘法拟合直线，返回描述已知数据点线性趋势的统计值。
LOGEST|known_y's;[known_x's];[const];[stats]|返回描述拟合已知数据点的指数曲线的统计值。
LOGINV|probability;mean;standard_dev|返回 x 的对数正态累积分布函数的反函数值，其中 ln(x) 服从以 Mean 和 Standard_dev 为参数的正态分布。
LOGNORM.DIST|x;mean;standard_dev;cumulative|返回 x 的对数正态分布，其中 ln(x) 服从以 Mean 和 Standard_dev 为参数的正态分布。
LOGNORM.INV|probability;mean;standard_dev|返回 x 的对数正态累积分布函数的反函数值。
LOGNORMDIST|x;mean;standard_dev|返回 x 的对数正态累积分布，其中 ln(x) 服从以 Mean 和 Standard_dev 为参数的正态分布。
MAX|number1;[number2];...|返回一组值中的最大值，忽略逻辑值和文本。
MAXA|value1;[value2];...|返回一组值中的最大值，不忽略逻辑值和文本。
MAXIFS|max_range;criteria_range1;criteria1;...|返回符合一组给定条件的单元格中的最大值。
MEDIAN|number1;[number2];...|返回中位数，即给定一组数字中居中的数字。
MIN|number1;[number2];...|返回一组值中的最小数字，忽略逻辑值和文本。
MINA|value1;[value2];...|返回一组值中的最小值，不忽略逻辑值和文本。
MINIFS|min_range;criteria_range1;criteria1;...|返回符合一组给定条件的单元格中的最小值。
MODE|number1;[number2];...|返回数组或数据区域中出现频率最高的值。
MODE.MULT|number1;[number2];...|以垂直数组形式返回数组或数据区域中出现频率最高的值。
MODE.SNGL|number1;[number2];...|返回数组或数据区域中出现频率最高的值。
NEGBINOM.DIST|number_f;number_s;probability_s;cumulative|返回负二项分布。
NEGBINOMDIST|number_f;number_s;probability_s|返回负二项分布，即第 Number_s 次成功之前出现 Number_f 次失败的概率。
NORM.DIST|x;mean;standard_dev;cumulative|返回指定平均值和标准偏差的正态分布。
NORM.INV|probability;mean;standard_dev|返回指定平均值和标准偏差的正态累积分布的反函数值。
NORM.S.DIST|z;cumulative|返回标准正态分布（平均值为 0，标准偏差为 1）。
NORM.S.INV|probability|返回标准正态累积分布的反函数值。
NORMDIST|x;mean;standard_dev;cumulative|返回指定平均值和标准偏差的正态累积分布。
NORMINV|probability;mean;standard_dev|返回指定平均值和标准偏差的正态累积分布的反函数值。
NORMSDIST|z|返回标准正态累积分布（平均值为 0，标准偏差为 1）。
NORMSINV|probability|返回标准正态累积分布（平均值为 0，标准偏差为 1）的反函数值。
PEARSON|array1;array2|返回 Pearson 积矩相关系数 r。
PERCENTILE|array;k|返回区域内数值的第 k 百分位数。
PERCENTILE.EXC|array;k|返回区域内数值的第 k 百分位数，其中 k 在 0 至 1 之间，不包括端点。
PERCENTILE.INC|array;k|返回区域内数值的第 k 百分位数，其中 k 在 0 至 1 之间，包括端点。
PERCENTRANK|array;x;[significance]|以百分比形式返回某个值在数据集中的排名。
PERCENTRANK.EXC|array;x;[significance]|以百分比形式返回某个值在数据集中的排名，范围为 0 至 1，不包括端点。
PERCENTRANK.INC|array;x;[significance]|以百分比形式返回某个值在数据集中的排名，范围为 0 至 1，包括端点。
PERMUT|number;number_chosen|返回从全部对象中选取指定数目对象的排列数。
PERMUTATIONA|number;number_chosen|返回从全部对象中选取指定数目对象的排列数，允许重复。
PHI|x|返回标准正态分布的密度函数值。
POISSON|x;mean;cumulative|返回 Poisson 分布。
POISSON.DIST|x;mean;cumulative|返回 Poisson 分布。
PROB|x_range;prob_range;lower_limit;[upper_limit]|返回区域内数值介于两个界限之间或等于下限的概率。
QUARTILE|array;quart|返回数据集的四分位数。
QUARTILE.EXC|array;quart|根据 0 至 1 之间（不包括端点）的百分位数，返回数据集的四分位数。
QUARTILE.INC|array;quart|根据 0 至 1 之间（包括端点）的百分位数，返回数据集的四分位数。
RANK|number;ref;[order]|返回数字在数值列表中的排名，即其大小相对于其他数值的位置。
RANK.AVG|number;ref;[order]|返回数字在数值列表中的排名；若多个值排名相同，则返回平均排名。
RANK.EQ|number;ref;[order]|返回数字在数值列表中的排名；若多个值排名相同，则返回这些值的最高排名。
RSQ|known_y's;known_x's|返回给定数据点的 Pearson 积矩相关系数的平方。
SKEW|number1;[number2];...|返回分布的偏度，用于描述分布围绕其平均值的不对称程度。
SKEW.P|number1;[number2];...|返回基于总体的分布偏度。
SLOPE|known_y's;known_x's|返回通过给定数据点的线性回归直线的斜率。
SMALL|array;k|返回数据集中第 k 个最小值，例如第五小的数字。
STANDARDIZE|x;mean;standard_dev|根据分布的平均值和标准偏差返回标准化值。
STDEV|number1;[number2];...|根据样本估算标准偏差，忽略样本中的逻辑值和文本。
STDEV.P|number1;[number2];...|计算以参数给出的整个总体的标准偏差，忽略逻辑值和文本。
STDEV.S|number1;[number2];...|根据样本估算标准偏差，忽略样本中的逻辑值和文本。
STDEVA|value1;[value2];...|根据样本估算标准偏差，包括逻辑值和文本。文本和 FALSE 视为 0，TRUE 视为 1。
STDEVP|number1;[number2];...|计算以参数给出的整个总体的标准偏差，忽略逻辑值和文本。
STDEVPA|value1;[value2];...|计算总体的标准偏差，包括逻辑值和文本。
STEYX|known_y's;known_x's|返回回归分析中各 x 对应的预测 y 值的标准误差。
T.DIST|x;deg_freedom;cumulative|返回左尾 Student t 分布。
T.DIST.2T|x;deg_freedom|返回双尾 Student t 分布。
T.DIST.RT|x;deg_freedom|返回右尾 Student t 分布。
T.INV|probability;deg_freedom|返回 Student t 分布的左尾反函数值。
T.INV.2T|probability;deg_freedom|返回 Student t 分布的双尾反函数值。
T.TEST|array1;array2;tails;type|返回 Student t 检验对应的概率。
TDIST|x;deg_freedom;tails|返回 Student t 分布。
TINV|probability;deg_freedom|返回 Student t 分布的双尾反函数值。
TREND|known_y's;[known_x's];[new_x's];[const]|使用最小二乘法，返回拟合已知数据点的线性趋势值。
TRIMMEAN|array;percent|返回去除两端部分数据后的平均值。
TTEST|array1;array2;tails;type|返回 Student t 检验对应的概率。
VAR|number1;[number2];...|根据样本估算方差，忽略样本中的逻辑值和文本。
VAR.P|number1;[number2];...|计算总体方差，忽略总体中的逻辑值和文本。
VAR.S|number1;[number2];...|根据样本估算方差，忽略样本中的逻辑值和文本。
VARA|value1;[value2];...|根据样本估算方差，包括逻辑值和文本。文本和 FALSE 视为 0，TRUE 视为 1。
VARP|number1;[number2];...|计算总体方差，忽略总体中的逻辑值和文本。
VARPA|value1;[value2];...|计算总体方差，包括逻辑值和文本。
WEIBULL|x;alpha;beta;cumulative|返回 Weibull 分布。
WEIBULL.DIST|x;alpha;beta;cumulative|返回 Weibull 分布。
Z.TEST|array;x;[sigma]|返回 z 检验的单尾 P 值。
ZTEST|array;x;[sigma]|Returns the one-tailed P-value of a z-test.`,
    'Lookup & Reference': `
ADDRESS|row_num;column_num;[abs_num];[a1];[sheet_text]|根据指定的行号和列号，以文本形式创建单元格引用。
AREAS|reference|返回引用中的区域数量。区域可以是一组连续单元格，也可以是单个单元格。
CHOOSE|index_num;value1;[value2];...|根据索引号，从值列表中选择一个值或要执行的操作。
CHOOSECOLS|array;col_num1;...|返回数组中的指定列。
CHOOSEROWS|array;row_num1;...|返回数组中的指定行。
COLUMN|[reference]|返回引用的列号。
COLUMNS|array|返回数组或引用的列数。
DROP|array;rows;[columns]|从数组开头或末尾删除指定的行或列。
EXPAND|array;rows;[columns];[pad_with]|将数组扩展到指定维度。
FILTER|array;include;[if_empty]|筛选区域或数组。
FORMULATEXT|reference|以字符串形式返回公式。
GETPIVOTDATA|data_field;pivot_table;[field1];[item1];...|提取数据透视表中存储的数据。
HLOOKUP|lookup_value;table_array;row_index_num;[range_lookup]|在表或数值数组的首行中查找值，然后返回同一列中指定行的值。
HSTACK|array1;[array2];...|按顺序水平拼接数组。
HYPERLINK|link_location;[friendly_name]|创建快捷方式或跳转链接，用于打开硬盘、网络服务器或互联网上的文档。
INDEX|array;row_num;[column_num];[area_num]|返回给定区域中指定行列交叉处单元格的值或引用。
INDIRECT|ref_text;[a1]|返回文本字符串指定的引用。
LOOKUP|lookup_value;lookup_vector;[result_vector]|从单行、单列区域或数组中查找值。
MATCH|lookup_value;lookup_array;[match_type]|按指定顺序匹配指定值，返回匹配项在数组中的相对位置。
OFFSET|reference;rows;cols;[height];[width]|返回相对于给定引用偏移指定行数和列数的区域引用。
ROW|[reference]|返回引用的行号。
ROWS|array|返回引用或数组的行数。
RTD|progID;server;topic1;...|从支持 COM 自动化的程序中检索实时数据。
SORT|array;[sort_index];[sort_order];[by_col]|对区域或数组排序。
SORTBY|array;by_array1;[sort_order1];...|根据对应区域或数组中的值，对区域或数组排序。
TAKE|array;rows;[columns]|返回数组开头或末尾的指定行或列。
TOCOL|array;[ignore];[scan_by_column]|将数组转换为单列返回。
TOROW|array;[ignore];[scan_by_column]|将数组转换为单行返回。
TRANSPOSE|array|将垂直单元格区域转换为水平区域，或反向转换。
UNIQUE|array;[by_col];[exactly_once]|返回区域或数组中的唯一值。
VLOOKUP|lookup_value;table_array;col_index_num;[range_lookup]|在表的最左列中查找值，然后返回同一行中指定列的值。
VSTACK|array1;[array2];...|按顺序垂直拼接数组。
WRAPCOLS|vector;wrap_count;[pad_with]|将一行或一列数值按指定元素数折行为多列。
WRAPROWS|vector;wrap_count;[pad_with]|将一行或一列数值按指定元素数折行为多行。
XLOOKUP|lookup_value;lookup_array;return_array;[if_not_found];[match_mode];[search_mode]|在区域或数组中查找匹配项，并返回第二个区域或数组中的对应项。
XMATCH|lookup_value;lookup_array;[match_mode];[search_mode]|Returns the relative position of an item in an array.`,
    Database: `
DAVERAGE|database;field;criteria|计算列表或数据库中符合指定条件的记录在某一列中的平均值。
DCOUNT|database;field;criteria|统计数据库中符合指定条件的记录在指定字段（列）中包含数字的单元格数。
DCOUNTA|database;field;criteria|统计数据库中符合指定条件的记录在指定字段（列）中的非空单元格数。
DGET|database;field;criteria|从数据库中提取符合指定条件的单条记录。
DMAX|database;field;criteria|返回数据库中符合指定条件的记录在指定字段（列）中的最大数字。
DMIN|database;field;criteria|返回数据库中符合指定条件的记录在指定字段（列）中的最小数字。
DPRODUCT|database;field;criteria|将数据库中符合指定条件的记录在指定字段（列）中的值相乘。
DSTDEV|database;field;criteria|根据选定数据库记录组成的样本估算标准偏差。
DSTDEVP|database;field;criteria|以选定的数据库记录为总体计算标准偏差。
DSUM|database;field;criteria|计算数据库中符合指定条件的记录的字段（列）数值之和。
DVAR|database;field;criteria|根据选定数据库记录组成的样本估算方差。
DVARP|database;field;criteria|Calculates variance based on the entire population of selected database entries.`,
    Text: `
ARRAYTOTEXT|array;[format]|返回数组的文本表示。
ASC|text|将全角（双字节）字符转换为半角（单字节）字符。
BAHTTEXT|number|将数字转换为泰铢文字。
CHAR|number|返回计算机字符集中指定数字代码对应的字符。
CLEAN|text|删除文本中的所有非打印字符。
CODE|text|按计算机使用的字符集，返回文本字符串首字符的数字代码。
CONCAT|text1;...|连接列表或区域中的文本字符串。
CONCATENATE|text1;[text2];...|将多个文本字符串连接为一个文本字符串。
DBCS|text|将半角（单字节）字符转换为全角（双字节）字符。
DOLLAR|number;[decimals]|使用货币格式将数字转换为文本。
EXACT|text1;text2|检查两个文本字符串是否完全相同，返回 TRUE 或 FALSE。EXACT 区分大小写。
FIND|find_text;within_text;[start_num]|返回一个文本字符串在另一个文本字符串中的起始位置。FIND 区分大小写。
FINDB|find_text;within_text;[start_num]|在一个文本字符串中查找另一个文本字符串的起始位置，按字节计数。
FIXED|number;[decimals];[no_commas]|将数字舍入到指定小数位数，以带或不带逗号的文本形式返回。
JIS|text|将半角（单字节）字符转换为全角（双字节）字符。
LEFT|text;[num_chars]|从文本字符串开头返回指定数量的字符。
LEFTB|text;[num_bytes]|从文本字符串开头返回指定数量的字节。
LEN|text|返回文本字符串的字符数。
LENB|text|返回文本字符串中各字符所占的字节数。
LOWER|text|将文本字符串中的所有字母转换为小写。
MID|text;start_num;num_chars|根据指定的起始位置和长度，从文本字符串中间返回字符。
MIDB|text;start_num;num_bytes|根据指定的起始位置和长度，从文本字符串中间返回字节。
NUMBERVALUE|text;[decimal_separator];[group_separator]|将文本转换为数字，不受区域设置影响。
PHONETIC|reference|从文本字符串中提取拼音（假名）字符。
PROPER|text|将文本字符串中每个单词的首字母转换为大写，其余字母转换为小写。
REPLACE|old_text;start_num;num_chars;new_text|将文本字符串的一部分替换为其他文本字符串。
REPLACEB|old_text;start_num;num_bytes;new_text|按字节计数，替换文本字符串的一部分。
REPT|text;number_times|将文本重复指定次数。使用 REPT 可将文本字符串重复填入单元格。
RIGHT|text;[num_chars]|从文本字符串末尾返回指定数量的字符。
RIGHTB|text;[num_bytes]|从文本字符串末尾返回指定数量的字节。
SEARCH|find_text;within_text;[start_num]|从左向右查找，返回指定字符或文本字符串首次出现的位置，不区分大小写。
SEARCHB|find_text;within_text;[start_num]|在一个文本值中查找另一个文本值，按字节计数，不区分大小写。
SUBSTITUTE|text;old_text;new_text;[instance_num]|将文本字符串中的现有文本替换为新文本。
T|value|检查值是否为文本；若是则返回该文本，否则返回双引号（空文本）。
TEXT|value;format_text|使用指定的数字格式将值转换为文本。
TEXTAFTER|text;delimiter;[instance_num];[match_mode];[match_end];[if_not_found]|返回指定字符或字符串之后的文本。
TEXTBEFORE|text;delimiter;[instance_num];[match_mode];[match_end];[if_not_found]|返回指定字符或字符串之前的文本。
TEXTJOIN|delimiter;ignore_empty;text1;...|使用分隔符连接列表或区域中的文本字符串。
TEXTSPLIT|text;col_delimiter;[row_delimiter];[ignore_empty];[match_mode];[pad_with]|使用分隔符将文本拆分为行或列。
TRIM|text|删除文本字符串中的多余空格，仅保留单词之间的单个空格。
UNICHAR|number|返回指定数值对应的 Unicode 字符。
UNICODE|text|返回文本首字符对应的数字（码位）。
UPPER|text|将文本字符串转换为全大写字母。
VALUE|text|将表示数字的文本字符串转换为数字。
VALUETOTEXT|value;[format]|Returns a text representation of a value.`,
    Logical: `
AND|logical1;[logical2];...|检查所有参数是否都为 TRUE；若是则返回 TRUE。
BYCOL|array;lambda|对每列应用 LAMBDA，并返回结果数组。
BYROW|array;lambda|对每行应用 LAMBDA，并返回结果数组。
FALSE||返回逻辑值 FALSE。
IF|logical_test;[value_if_true];[value_if_false]|检查条件是否满足；若为 TRUE 则返回一个值，若为 FALSE 则返回另一个值。
IFERROR|value;value_if_error|若表达式为错误，则返回 value_if_error，否则返回表达式本身的值。
IFNA|value;value_if_na|若表达式结果为 #N/A，则返回指定值，否则返回表达式的结果。
IFS|logical_test1;value_if_true1;...|检查一个或多个条件是否满足，返回第一个 TRUE 条件对应的值。
LAMBDA|[parameter1];...;calculation|创建可在公式中调用的函数值。
LET|name1;name_value1;...;calculation|为名称指定计算值，以便保存中间计算结果。
MAKEARRAY|rows;columns;lambda|通过应用 LAMBDA，返回指定行数和列数的数组。
MAP|array1;...;lambda|将 LAMBDA 应用于数组中的每个值，返回映射后的新值组成的数组。
NOT|logical|将 FALSE 转为 TRUE，或将 TRUE 转为 FALSE。
OR|logical1;[logical2];...|检查是否有参数为 TRUE，返回 TRUE 或 FALSE。仅在所有参数都为 FALSE 时返回 FALSE。
REDUCE|[initial_value];array;lambda|将 LAMBDA 应用于数组中的每个值，返回累积计算的结果。
SCAN|[initial_value];array;lambda|将 LAMBDA 依次应用于数组中的每个值，返回包含每个中间结果的数组。
SWITCH|expression;value1;result1;...|将表达式与值列表比较，返回第一个匹配值对应的结果。
TRUE||返回逻辑值 TRUE。
XOR|logical1;[logical2];...|Returns a logical 'Exclusive Or' of all arguments.`,
    Information: `
CELL|info_type;[reference]|按工作表的阅读顺序，返回引用中第一个单元格的格式、位置或内容信息。
ERROR.TYPE|error_val|返回与错误值对应的数字。
INFO|type_text|返回当前操作环境的信息。
ISBLANK|value|检查引用是否指向空单元格，返回 TRUE 或 FALSE。
ISERR|value|检查值是否为 #N/A 以外的错误，返回 TRUE 或 FALSE。
ISERROR|value|检查值是否为错误，返回 TRUE 或 FALSE。
ISEVEN|number|若数字为偶数，则返回 TRUE。
ISFORMULA|reference|检查引用的单元格是否包含公式，返回 TRUE 或 FALSE。
ISLOGICAL|value|检查值是否为逻辑值（TRUE 或 FALSE），返回 TRUE 或 FALSE。
ISNA|value|检查值是否为 #N/A，返回 TRUE 或 FALSE。
ISNONTEXT|value|检查值是否不是文本（空单元格不是文本），返回 TRUE 或 FALSE。
ISNUMBER|value|检查值是否为数字，返回 TRUE 或 FALSE。
ISODD|number|若数字为奇数，则返回 TRUE。
ISOMITTED|argument|检查 LAMBDA 中是否缺少参数值，返回 TRUE 或 FALSE。
ISREF|value|检查值是否为引用，返回 TRUE 或 FALSE。
ISTEXT|value|检查值是否为文本，返回 TRUE 或 FALSE。
N|value|将非数值转换为数字：日期转换为序列号，TRUE 转换为 1，其他值转换为 0（零）。
NA||返回错误值 #N/A（值不可用）。
SHEET|[value]|返回所引用工作表的编号。
SHEETS|[reference]|返回引用中的工作表数量。
TYPE|value|Returns an integer representing the data type of a value: number = 1; text = 2; logical value = 4; error value = 16; array = 64.`,
    Engineering: `
BESSELI|x;n|返回修正贝塞尔函数 In(x)。
BESSELJ|x;n|返回贝塞尔函数 Jn(x)。
BESSELK|x;n|返回修正贝塞尔函数 Kn(x)。
BESSELY|x;n|返回贝塞尔函数 Yn(x)。
BIN2DEC|number|将二进制数转换为十进制数。
BIN2HEX|number;[places]|将二进制数转换为十六进制数。
BIN2OCT|number;[places]|将二进制数转换为八进制数。
BITAND|number1;number2|返回两个数字的按位与运算结果。
BITLSHIFT|number;shift_amount|返回将数字向左移动 shift_amount 位的结果。
BITOR|number1;number2|返回两个数字的按位或运算结果。
BITRSHIFT|number;shift_amount|返回将数字向右移动 shift_amount 位的结果。
BITXOR|number1;number2|返回两个数字的按位异或运算结果。
COMPLEX|real_num;i_num;[suffix]|将实部和虚部系数组合成复数。
CONVERT|number;from_unit;to_unit|将数值从一种计量单位转换为另一种计量单位。
DEC2BIN|number;[places]|将十进制数转换为二进制数。
DEC2HEX|number;[places]|将十进制数转换为十六进制数。
DEC2OCT|number;[places]|将十进制数转换为八进制数。
DELTA|number1;[number2]|测试两个数字是否相等。
ERF|lower_limit;[upper_limit]|返回误差函数。
ERF.PRECISE|x|返回误差函数。
ERFC|x|返回互补误差函数。
ERFC.PRECISE|x|返回互补误差函数。
GESTEP|number;[step]|测试数字是否大于阈值。
HEX2BIN|number;[places]|将十六进制数转换为二进制数。
HEX2DEC|number|将十六进制数转换为十进制数。
HEX2OCT|number;[places]|将十六进制数转换为八进制数。
IMABS|inumber|返回复数的绝对值（模）。
IMAGINARY|inumber|返回复数的虚部系数。
IMARGUMENT|inumber|返回辐角 q，以弧度表示。
IMCONJUGATE|inumber|返回复数的共轭复数。
IMCOS|inumber|返回复数的余弦值。
IMCOSH|inumber|返回复数的双曲余弦值。
IMCOT|inumber|返回复数的余切值。
IMCSC|inumber|返回复数的余割值。
IMCSCH|inumber|返回复数的双曲余割值。
IMDIV|inumber1;inumber2|返回两个复数的商。
IMEXP|inumber|返回复数的指数值。
IMLN|inumber|返回复数的自然对数。
IMLOG10|inumber|返回复数以 10 为底的对数。
IMLOG2|inumber|返回复数以 2 为底的对数。
IMPOWER|inumber;number|返回复数的整数次幂。
IMPRODUCT|inumber1;[inumber2];...|返回 1 至 255 个复数的乘积。
IMREAL|inumber|返回复数的实部系数。
IMSEC|inumber|返回复数的正割值。
IMSECH|inumber|返回复数的双曲正割值。
IMSIN|inumber|返回复数的正弦值。
IMSINH|inumber|返回复数的双曲正弦值。
IMSQRT|inumber|返回复数的平方根。
IMSUB|inumber1;inumber2|返回两个复数的差。
IMSUM|inumber1;[inumber2];...|返回复数之和。
IMTAN|inumber|返回复数的正切值。
OCT2BIN|number;[places]|将八进制数转换为二进制数。
OCT2DEC|number|将八进制数转换为十进制数。
OCT2HEX|number;[places]|Converts an octal number to hexadecimal.`,
    Cube: `
CUBEKPIMEMBER|connection;kpi_name;kpi_property;[caption]|返回关键绩效指标 (KPI) 的属性，并在单元格中显示 KPI 名称。
CUBEMEMBER|connection;member_expression;[caption]|返回多维数据集中的成员或元组。
CUBEMEMBERPROPERTY|connection;member_expression;property|返回多维数据集中成员属性的值。
CUBERANKEDMEMBER|connection;set_expression;rank;[caption]|返回集合中排名第 n 的成员。
CUBESET|connection;set_expression;[caption];[sort_order];[sort_by]|向服务器上的多维数据集发送集合表达式，以定义由成员或元组组成的计算集合。
CUBESETCOUNT|set|返回集合中的项目数。
CUBEVALUE|connection;[member_expression1];...|Returns an aggregated value from the cube.`,
    Web: `
ENCODEURL|text|返回经过 URL 编码的字符串。
FILTERXML|xml;xpath|使用指定的 XPath，从 XML 内容中返回特定数据。
WEBSERVICE|url|Returns data from a web service.`,
  };
  const OLD = new Set('ACCRINT ACCRINTM AMORDEGRC AMORLINC COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD CUMIPMT CUMPRINC DB DDB DISC DOLLARDE DOLLARFR DURATION EFFECT FV FVSCHEDULE INTRATE IPMT IRR ISPMT MDURATION MIRR NOMINAL NPER NPV ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD PMT PPMT PRICE PRICEDISC PRICEMAT PV RATE RECEIVED SLN SYD TBILLEQ TBILLPRICE TBILLYIELD VDB XIRR XNPV YIELD YIELDDISC YIELDMAT DATE DATEVALUE DAY DAYS360 EDATE EOMONTH HOUR MINUTE MONTH NETWORKDAYS NOW SECOND TIME TIMEVALUE TODAY WEEKDAY WEEKNUM WORKDAY YEAR YEARFRAC ABS ACOS ACOSH ASIN ASINH ATAN ATAN2 ATANH CEILING COMBIN COS COSH DEGREES EVEN EXP FACT FACTDOUBLE FLOOR GCD INT LCM LN LOG LOG10 MDETERM MINVERSE MMULT MOD MROUND MULTINOMIAL ODD PI POWER PRODUCT QUOTIENT RADIANS RAND RANDBETWEEN ROMAN ROUND ROUNDDOWN ROUNDUP SERIESSUM SIGN SIN SINH SQRT SQRTPI SUBTOTAL SUM SUMIF SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 TAN TANH TRUNC AVEDEV AVERAGE AVERAGEA BETADIST BETAINV BINOMDIST CHIDIST CHIINV CHITEST CONFIDENCE CORREL COUNT COUNTA COUNTBLANK COUNTIF COVAR CRITBINOM DEVSQ EXPONDIST FDIST FINV FISHER FISHERINV FORECAST FREQUENCY FTEST GAMMADIST GAMMAINV GAMMALN GEOMEAN GROWTH HARMEAN HYPGEOMDIST INTERCEPT KURT LARGE LINEST LOGEST LOGINV LOGNORMDIST MAX MAXA MEDIAN MIN MINA MODE NEGBINOMDIST NORMDIST NORMINV NORMSDIST NORMSINV PEARSON PERCENTILE PERCENTRANK PERMUT POISSON PROB QUARTILE RANK RSQ SKEW SLOPE SMALL STANDARDIZE STDEV STDEVA STDEVP STDEVPA STEYX TDIST TINV TREND TRIMMEAN TTEST VAR VARA VARP VARPA WEIBULL ZTEST ADDRESS AREAS CHOOSE COLUMN COLUMNS GETPIVOTDATA HLOOKUP HYPERLINK INDEX INDIRECT LOOKUP MATCH OFFSET ROW ROWS RTD TRANSPOSE VLOOKUP DAVERAGE DCOUNT DCOUNTA DGET DMAX DMIN DPRODUCT DSTDEV DSTDEVP DSUM DVAR DVARP ASC BAHTTEXT CHAR CLEAN CODE CONCATENATE DOLLAR EXACT FIND FINDB FIXED JIS LEFT LEFTB LEN LENB LOWER MID MIDB PHONETIC PROPER REPLACE REPLACEB REPT RIGHT RIGHTB SEARCH SEARCHB SUBSTITUTE T TEXT TRIM UPPER VALUE AND FALSE IF NOT OR TRUE CELL ERROR.TYPE INFO ISBLANK ISERR ISERROR ISEVEN ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISREF ISTEXT N NA TYPE BESSELI BESSELJ BESSELK BESSELY BIN2DEC BIN2HEX BIN2OCT COMPLEX CONVERT DEC2BIN DEC2HEX DEC2OCT DELTA ERF ERFC GESTEP HEX2BIN HEX2DEC HEX2OCT IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMDIV IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSIN IMSQRT IMSUB IMSUM OCT2BIN OCT2DEC OCT2HEX'.split(' '));
  const map = new Map();
  for (const [cat, text] of Object.entries(DATA)) {
    for (const line of text.trim().split('\n')) {
      const [name, args, desc] = line.split('|');
      map.set(name, { name, cat, args: args ? args.split(';') : [], desc: desc || '', classic: OLD.has(name) });
    }
  }
  FI.get = (name) => map.get(String(name).toUpperCase()) || null;
  FI.all = () => Array.from(map.values());
  /* the categories are ids (f.cat, Insert Function's 'All' and 'Most Recently Used') shown translated */
  FI.categories = [['Most Recently Used', '最近使用'], ['All', '全部'], ['Financial', '财务'], ['Date & Time', '日期和时间'], ['Math & Trig', '数学与三角函数'], ['Statistical', '统计'], ['Lookup & Reference', '查找与引用'], ['Database', '数据库'], ['Text', '文本'], ['Logical', '逻辑'], ['Information', '信息'], ['Engineering', '工程'], ['Cube', '立方体'], ['Web', 'Web']];
  FI.syntax = (f) => f.name + '(' + f.args.join(',') + ')';
  /* argument help shown in the Function Arguments dialog (translated there) */
  const ARGHELP = {
    number: L.N('数字，或包含数字的单元格引用。'), number1: L.N('number1、number2、... 是要处理的 1 至 255 个数字或参数。'),
    value1: L.N('value1、value2、... 是 1 至 255 个值或区域。'), logical1: L.N('logical1、logical2、... 是要检测的 1 至 255 个条件，每个条件可以为 TRUE 或 FALSE。'),
    logical_test: L.N('可求值为 TRUE 或 FALSE 的任意值或表达式。'), value_if_true: L.N('Logical_test 为 TRUE 时返回的值。省略时返回 TRUE。'), value_if_false: L.N('Logical_test 为 FALSE 时返回的值。省略时返回 FALSE。'),
    lookup_value: L.N('要在表的首列（或首行）中查找的值，可以是数值、引用或文本字符串。'), table_array: L.N('要从中检索数据的表，其中包含文本、数字或逻辑值。Table_array 可以是区域引用或区域名称。'),
    col_index_num: L.N('table_array 中要返回匹配值的列号。表中的第一列为第 1 列。'), row_index_num: L.N('table_array 中要返回匹配值的行号。'),
    range_lookup: L.N('逻辑值：在首列（按升序排列）中查找近似匹配为 TRUE 或省略；查找精确匹配为 FALSE。'),
    range: L.N('要计算的单元格区域。'), criteria: L.N('以数字、表达式或文本形式给出的条件，用于确定要相加的单元格。'), sum_range: L.N('是实际要求和的单元格。省略时使用 range 中的单元格。'),
    text: L.N('要处理的文本字符串。'), num_chars: L.N('指定所需字符数。'), rate: L.N('每期利率。例如，年利率为 6%、按季度付款时，使用 6%/4。'),
    nper: L.N('投资的总付款期数。'), pv: L.N('现值，即未来一系列付款在当前的总价值。'), fv: L.N('终值，即最后一次付款后希望达到的现金余额。'), type: L.N('逻辑值：期初付款为 1，期末付款为 0 或省略。'),
    num_digits: L.N('要舍入到的位数。负数表示小数点左侧的位数，零表示舍入到最接近的整数。'), serial_number: L.N('Ledger 日期时间代码中的数字。'),
    array: L.N('单元格区域或数组常量。'), k: L.N('要返回的数据在数组或单元格区域中的排名（从最大或最小值开始）。'), reference: L.N('单元格或单元格区域。'),
    database: L.N('构成列表或数据库的单元格区域。数据库是相关数据组成的列表。'), field: L.N('用双引号括起的列标签，或表示列在列表中位置的数字。'), criteria_range1: L.N('要按指定条件判断的单元格区域。'),
  };
  FI.argHelp = (name) => ARGHELP[String(name).replace(/[[\]]|\.\.\./g, '')] || '';
})(typeof window !== 'undefined' ? window : globalThis);
