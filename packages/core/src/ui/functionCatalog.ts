/**
 * Function signatures and short descriptions for formula autocomplete, argument hints and the
 * Insert Function dialog. Functions the engine implements but this table does not describe get a
 * generic signature derived from their arity.
 */
import { registry } from '../core/formula/functions';

export interface FunctionInfo {
  name: string;
  category: string;
  /** Parameters; optional ones are wrapped in brackets, a trailing '…' marks repetition. */
  args: string[];
  description: string;
}

const CATEGORIES: Record<string, string> = {
  M: 'Math & trigonometry',
  S: 'Statistical',
  T: 'Text',
  L: 'Logical',
  R: 'Lookup & reference',
  D: 'Date & time',
  F: 'Financial',
  E: 'Engineering',
  I: 'Information',
  B: 'Database',
  W: 'Web',
};

// NAME|category|arguments separated by ';'|description
const TABLE = `
ABS|M|number|Absolute value of a number.
ACCRINT|F|issue;first_interest;settlement;rate;par;frequency;[basis];[calc_method]|Accrued interest of a security that pays periodic interest.
ACCRINTM|F|issue;settlement;rate;par;[basis]|Accrued interest of a security that pays interest at maturity.
ACOS|M|number|Arccosine of a number, in radians.
ACOSH|M|number|Inverse hyperbolic cosine.
ACOT|M|number|Arccotangent of a number, in radians.
ACOTH|M|number|Inverse hyperbolic cotangent.
ADDRESS|R|row_num;column_num;[abs_num];[a1];[sheet_text]|Builds a cell reference as text.
AGGREGATE|M|function_num;options;ref1;[ref2];…|Aggregate (sum, average, …) that can skip hidden rows and errors.
AMORDEGRC|F|cost;date_purchased;first_period;salvage;period;rate;[basis]|Depreciation of an accounting period with a depreciation coefficient (French system).
AMORLINC|F|cost;date_purchased;first_period;salvage;period;rate;[basis]|Depreciation of an accounting period (French system).
AND|L|logical1;[logical2];…|TRUE when every argument is TRUE.
ARABIC|M|text|Converts a Roman numeral to a number.
AREAS|R|reference|Number of areas in a reference.
ARRAYTOTEXT|T|array;[format]|Text representation of an array.
ASC|T|text|Converts full-width characters to half-width ones.
ASIN|M|number|Arcsine of a number, in radians.
ASINH|M|number|Inverse hyperbolic sine.
ATAN|M|number|Arctangent of a number, in radians.
ATAN2|M|x_num;y_num|Arctangent of the point (x, y).
ATANH|M|number|Inverse hyperbolic tangent.
AVEDEV|S|number1;[number2];…|Average absolute deviation from the mean.
AVERAGE|S|number1;[number2];…|Arithmetic mean of the numbers.
AVERAGEA|S|value1;[value2];…|Mean that counts text as 0 and TRUE as 1.
AVERAGEIF|S|range;criteria;[average_range]|Mean of the cells that meet a condition.
AVERAGEIFS|S|average_range;criteria_range1;criteria1;[criteria_range2;criteria2];…|Mean of the cells that meet several conditions.
BAHTTEXT|T|number|Spells a number in Thai words with the baht currency.
BASE|M|number;radix;[min_length]|Converts a number to text in another base.
BESSELI|E|x;n|Modified Bessel function In(x).
BESSELJ|E|x;n|Bessel function Jn(x).
BESSELK|E|x;n|Modified Bessel function Kn(x).
BESSELY|E|x;n|Bessel function Yn(x).
BETA.DIST|S|x;alpha;beta;cumulative;[A];[B]|Beta distribution.
BETA.INV|S|probability;alpha;beta;[A];[B]|Inverse of the cumulative beta distribution.
BETADIST|S|x;alpha;beta;[A];[B]|Cumulative beta distribution (compatibility).
BETAINV|S|probability;alpha;beta;[A];[B]|Inverse of the cumulative beta distribution (compatibility).
BIN2DEC|E|number|Binary to decimal.
BIN2HEX|E|number;[places]|Binary to hexadecimal.
BIN2OCT|E|number;[places]|Converts a binary number to octal.
BINOM.DIST|S|number_s;trials;probability_s;cumulative|Binomial distribution probability.
BINOM.DIST.RANGE|S|trials;probability_s;number_s;[number_s2]|Probability of a number (or a range) of successes in trials.
BINOM.INV|S|trials;probability_s;alpha|Smallest number of successes whose cumulative binomial probability reaches alpha.
BINOMDIST|S|number_s;trials;probability_s;cumulative|Binomial distribution probability (compatibility).
BITAND|E|number1;number2|Bitwise AND of two numbers.
BITLSHIFT|E|number;shift_amount|Shifts the bits of a number to the left.
BITOR|E|number1;number2|Bitwise OR of two numbers.
BITRSHIFT|E|number;shift_amount|Shifts the bits of a number to the right.
BITXOR|E|number1;number2|Bitwise exclusive OR of two numbers.
BYCOL|L|array;lambda|Applies a LAMBDA to each column and returns the results.
BYROW|L|array;lambda|Applies a LAMBDA to each row and returns the results.
CEILING|M|number;significance|Rounds up to a multiple of significance.
CEILING.MATH|M|number;[significance];[mode]|Rounds up to a multiple of significance.
CEILING.PRECISE|M|number;[significance]|Rounds a number up to a multiple, whatever its sign.
CELL|I|info_type;[reference]|Information about a cell's formatting, location or contents.
CHAR|T|number|Character for a code number.
CHIDIST|S|x;deg_freedom|Right-tailed chi-squared probability (compatibility).
CHIINV|S|probability;deg_freedom|Inverse of the right-tailed chi-squared distribution (compatibility).
CHISQ.DIST|S|x;deg_freedom;cumulative|Chi-squared distribution.
CHISQ.DIST.RT|S|x;deg_freedom|Right-tailed chi-squared probability.
CHISQ.INV|S|probability;deg_freedom|Inverse of the left-tailed chi-squared distribution.
CHISQ.INV.RT|S|probability;deg_freedom|Inverse of the right-tailed chi-squared distribution.
CHISQ.TEST|S|actual_range;expected_range|Chi-squared test for independence.
CHITEST|S|actual_range;expected_range|Chi-squared test for independence (compatibility).
CHOOSE|R|index_num;value1;[value2];…|Picks a value from the list by position.
CHOOSECOLS|R|array;col_num1;[col_num2];…|Returns the given columns of an array.
CHOOSEROWS|R|array;row_num1;[row_num2];…|Returns the given rows of an array.
CLEAN|T|text|Removes non-printable characters.
CODE|T|text|Code number of the first character.
COLUMN|R|[reference]|Column number of a reference.
COLUMNS|R|array|Number of columns in a reference or array.
COMBIN|M|number;number_chosen|Number of combinations.
COMBINA|M|number;number_chosen|Number of combinations with repetitions.
COMPLEX|E|real_num;i_num;[suffix]|Builds a complex number from its real and imaginary parts.
CONCAT|T|text1;[text2];…|Joins texts and ranges.
CONCATENATE|T|text1;[text2];…|Joins texts.
CONFIDENCE|S|alpha;standard_dev;size|Confidence interval of a population mean (compatibility).
CONFIDENCE.NORM|S|alpha;standard_dev;size|Confidence interval of a population mean, normal distribution.
CONFIDENCE.T|S|alpha;standard_dev;size|Confidence interval of a population mean, Student's t distribution.
CONVERT|E|number;from_unit;to_unit|Converts between measurement units.
CORREL|S|array1;array2|Correlation coefficient of two data sets.
COS|M|number|Cosine of an angle in radians.
COSH|M|number|Hyperbolic cosine.
COT|M|number|Cotangent of an angle.
COTH|M|number|Hyperbolic cotangent.
COUNT|S|value1;[value2];…|Counts the numbers.
COUNTA|S|value1;[value2];…|Counts the non-empty values.
COUNTBLANK|S|range|Counts empty cells.
COUNTIF|S|range;criteria|Counts the cells that meet a condition.
COUNTIFS|S|criteria_range1;criteria1;[criteria_range2;criteria2];…|Counts the cells that meet several conditions.
COUPDAYBS|F|settlement;maturity;frequency;[basis]|Days from the start of the coupon period to the settlement date.
COUPDAYS|F|settlement;maturity;frequency;[basis]|Days in the coupon period that contains the settlement date.
COUPDAYSNC|F|settlement;maturity;frequency;[basis]|Days from the settlement date to the next coupon date.
COUPNCD|F|settlement;maturity;frequency;[basis]|Next coupon date after the settlement date.
COUPNUM|F|settlement;maturity;frequency;[basis]|Number of coupons payable between settlement and maturity.
COUPPCD|F|settlement;maturity;frequency;[basis]|Last coupon date before the settlement date.
COVAR|S|array1;array2|Population covariance (compatibility).
COVARIANCE.P|S|array1;array2|Population covariance.
COVARIANCE.S|S|array1;array2|Sample covariance.
CRITBINOM|S|trials;probability_s;alpha|Smallest value whose cumulative binomial distribution reaches alpha (compatibility).
CSC|M|number|Cosecant of an angle.
CSCH|M|number|Hyperbolic cosecant.
CUMIPMT|F|rate;nper;pv;start_period;end_period;type|Cumulative interest paid between two periods.
CUMPRINC|F|rate;nper;pv;start_period;end_period;type|Cumulative principal paid between two periods.
DATE|D|year;month;day|Serial number of a date.
DATEDIF|D|start_date;end_date;unit|Difference between two dates in days, months or years.
DATEVALUE|D|date_text|Converts date text to a serial number.
DAVERAGE|B|database;field;criteria|Average of matching database records.
DAY|D|serial_number|Day of the month (1-31).
DAYS|D|end_date;start_date|Number of days between two dates.
DAYS360|D|start_date;end_date;[method]|Days between dates on a 360-day year.
DB|F|cost;salvage;life;period;[month]|Fixed-declining balance depreciation.
DBCS|T|text|Converts half-width characters to full-width ones.
DCOUNT|B|database;field;criteria|Counts numeric cells in matching records.
DCOUNTA|B|database;field;criteria|Counts non-empty cells in matching records.
DDB|F|cost;salvage;life;period;[factor]|Double-declining balance depreciation.
DEC2BIN|E|number;[places]|Decimal to binary.
DEC2HEX|E|number;[places]|Decimal to hexadecimal.
DEC2OCT|E|number;[places]|Converts a decimal number to octal.
DECIMAL|M|text;radix|Converts text in another base to a number.
DEGREES|M|angle|Radians to degrees.
DELTA|E|number1;[number2]|1 when two numbers are equal, else 0.
DEVSQ|S|number1;[number2];…|Sum of squared deviations from the mean.
DGET|B|database;field;criteria|Single value from the matching record.
DISC|F|settlement;maturity;pr;redemption;[basis]|Discount rate of a security.
DMAX|B|database;field;criteria|Largest value in matching records.
DMIN|B|database;field;criteria|Smallest value in matching records.
DOLLAR|T|number;[decimals]|Formats a number as currency text.
DOLLARDE|F|fractional_dollar;fraction|Converts a dollar price written as a fraction into a decimal number.
DOLLARFR|F|decimal_dollar;fraction|Writes a decimal dollar price as a fraction.
DPRODUCT|B|database;field;criteria|Product of a field over the matching records.
DROP|R|array;rows;[columns]|Removes rows or columns from the start or end of an array.
DSTDEV|B|database;field;criteria|Sample standard deviation of a field over the matching records.
DSTDEVP|B|database;field;criteria|Population standard deviation of a field over the matching records.
DSUM|B|database;field;criteria|Sum of matching records.
DURATION|F|settlement;maturity;coupon;yld;frequency;[basis]|Macaulay duration of a security.
DVAR|B|database;field;criteria|Sample variance of a field over the matching records.
DVARP|B|database;field;criteria|Population variance of a field over the matching records.
EDATE|D|start_date;months|Date a number of months before or after a date.
EFFECT|F|nominal_rate;npery|Effective annual interest rate.
ENCODEURL|W|text|Text encoded for use in a URL.
EOMONTH|D|start_date;months|Last day of the month, months away.
ERF|E|lower_limit;[upper_limit]|Error function.
ERF.PRECISE|E|x|Error function between 0 and x.
ERFC|E|x|Complementary error function.
ERFC.PRECISE|E|x|Complementary error function between x and infinity.
ERROR.TYPE|I|error_val|Number that identifies an error value.
EUROCONVERT|F|number;source;target;[full_precision];[triangulation_precision]|Converts between the euro and the former currencies of the euro area.
EVEN|M|number|Rounds up to the nearest even integer.
EXACT|T|text1;text2|TRUE when two texts are identical (case-sensitive).
EXP|M|number|e raised to a power.
EXPAND|R|array;rows;[columns];[pad_with]|Expands an array to the given size.
EXPON.DIST|S|x;lambda;cumulative|Exponential distribution.
EXPONDIST|S|x;lambda;cumulative|Exponential distribution (compatibility).
F.DIST|S|x;deg_freedom1;deg_freedom2;cumulative|F probability distribution.
F.DIST.RT|S|x;deg_freedom1;deg_freedom2|Right-tailed F probability distribution.
F.INV|S|probability;deg_freedom1;deg_freedom2|Inverse of the F probability distribution.
F.INV.RT|S|probability;deg_freedom1;deg_freedom2|Inverse of the right-tailed F probability distribution.
F.TEST|S|array1;array2|F-test: probability that the variances of two arrays do not differ significantly.
FACT|M|number|Factorial.
FACTDOUBLE|M|number|Double factorial of a number.
FALSE|L||The logical value FALSE.
FDIST|S|x;deg_freedom1;deg_freedom2|Right-tailed F probability distribution (compatibility).
FILTER|R|array;include;[if_empty]|Rows of an array that meet a condition.
FILTERXML|W|xml;xpath|Values selected from XML text by an XPath expression.
FIND|T|find_text;within_text;[start_num]|Position of text within text (case-sensitive).
FINDB|T|find_text;within_text;[start_num]|Position of a text within another, counting bytes.
FINV|S|probability;deg_freedom1;deg_freedom2|Inverse of the right-tailed F distribution (compatibility).
FISHER|S|x|Fisher transformation.
FISHERINV|S|y|Inverse of the Fisher transformation.
FIXED|T|number;[decimals];[no_commas]|Formats a number as text with fixed decimals.
FLOOR|M|number;significance|Rounds down to a multiple of significance.
FLOOR.MATH|M|number;[significance];[mode]|Rounds down to a multiple of significance.
FLOOR.PRECISE|M|number;[significance]|Rounds a number down to a multiple, whatever its sign.
FORECAST|S|x;known_ys;known_xs|Linear-trend prediction.
FORECAST.LINEAR|S|x;known_ys;known_xs|Linear-trend prediction.
FORMULATEXT|R|reference|Formula of a cell as text.
FREQUENCY|S|data_array;bins_array|Frequency distribution as a vertical array.
FTEST|S|array1;array2|F-test (compatibility).
FV|F|rate;nper;pmt;[pv];[type]|Future value of an investment.
FVSCHEDULE|F|principal;schedule|Future value of a principal after a series of interest rates.
GAMMA|S|x|Value of the gamma function.
GAMMA.DIST|S|x;alpha;beta;cumulative|Gamma distribution.
GAMMA.INV|S|probability;alpha;beta|Inverse of the cumulative gamma distribution.
GAMMADIST|S|x;alpha;beta;cumulative|Gamma distribution (compatibility).
GAMMAINV|S|probability;alpha;beta|Inverse of the cumulative gamma distribution (compatibility).
GAMMALN|S|x|Natural logarithm of the gamma function.
GAMMALN.PRECISE|S|x|Natural logarithm of the gamma function.
GAUSS|S|z|Probability that a standard normal value lies between the mean and z standard deviations from it.
GCD|M|number1;[number2];…|Greatest common divisor.
GEOMEAN|S|number1;[number2];…|Geometric mean.
GESTEP|E|number;[step]|1 when a number is at least the step, else 0.
GROUPBY|R|row_fields;values;function;[field_headers];[total_depth];[sort_order];[filter_array];[field_relationship]|Groups values by row fields and aggregates them.
GROWTH|S|known_ys;[known_xs];[new_xs];[const]|Exponential-trend prediction.
HARMEAN|S|number1;[number2];…|Harmonic mean.
HEX2BIN|E|number;[places]|Converts a hexadecimal number to binary.
HEX2DEC|E|number|Hexadecimal to decimal.
HEX2OCT|E|number;[places]|Converts a hexadecimal number to octal.
HLOOKUP|R|lookup_value;table_array;row_index_num;[range_lookup]|Looks along the top row and returns a value from another row.
HOUR|D|serial_number|Hour (0-23).
HSTACK|R|array1;[array2];…|Stacks arrays side by side.
HYPERLINK|I|link_location;[friendly_name]|Creates a link.
HYPGEOM.DIST|S|sample_s;number_sample;population_s;number_pop;cumulative|Hypergeometric distribution.
HYPGEOMDIST|S|sample_s;number_sample;population_s;number_pop|Hypergeometric distribution (compatibility).
IF|L|logical_test;[value_if_true];[value_if_false]|Returns one value when a condition is TRUE and another when it is FALSE.
IFERROR|L|value;value_if_error|Value, or an alternative when it is an error.
IFNA|L|value;value_if_na|Value, or an alternative when it is #N/A.
IFS|L|logical_test1;value_if_true1;[logical_test2;value_if_true2];…|Value of the first condition that is TRUE.
IMABS|E|inumber|Absolute value (modulus) of a complex number.
IMAGINARY|E|inumber|Imaginary coefficient of a complex number.
IMARGUMENT|E|inumber|Argument (angle in radians) of a complex number.
IMCONJUGATE|E|inumber|Complex conjugate of a complex number.
IMCOS|E|inumber|Cosine of a complex number.
IMCOSH|E|inumber|Hyperbolic cosine of a complex number.
IMCOT|E|inumber|Cotangent of a complex number.
IMCSC|E|inumber|Cosecant of a complex number.
IMCSCH|E|inumber|Hyperbolic cosecant of a complex number.
IMDIV|E|inumber1;inumber2|Quotient of two complex numbers.
IMEXP|E|inumber|Exponential of a complex number.
IMLN|E|inumber|Natural logarithm of a complex number.
IMLOG10|E|inumber|Base-10 logarithm of a complex number.
IMLOG2|E|inumber|Base-2 logarithm of a complex number.
IMPOWER|E|inumber;number|Complex number raised to a power.
IMPRODUCT|E|inumber1;[inumber2];…|Product of complex numbers.
IMREAL|E|inumber|Real coefficient of a complex number.
IMSEC|E|inumber|Secant of a complex number.
IMSECH|E|inumber|Hyperbolic secant of a complex number.
IMSIN|E|inumber|Sine of a complex number.
IMSINH|E|inumber|Hyperbolic sine of a complex number.
IMSQRT|E|inumber|Square root of a complex number.
IMSUB|E|inumber1;inumber2|Difference of two complex numbers.
IMSUM|E|inumber1;[inumber2];…|Sum of complex numbers.
IMTAN|E|inumber|Tangent of a complex number.
INDEX|R|array;row_num;[column_num];[area_num]|Value at a row and column of a range.
INDIRECT|R|ref_text;[a1]|Reference given as text.
INFO|I|type_text|Information about the current operating environment.
INT|M|number|Rounds down to an integer.
INTERCEPT|S|known_ys;known_xs|Y-intercept of the regression line.
INTRATE|F|settlement;maturity;investment;redemption;[basis]|Interest rate of a fully invested security.
IPMT|F|rate;per;nper;pv;[fv];[type]|Interest part of a payment.
IRR|F|values;[guess]|Internal rate of return.
ISBLANK|I|value|TRUE when the cell is empty.
ISERR|I|value|TRUE for any error except #N/A.
ISERROR|I|value|TRUE for any error.
ISEVEN|I|number|TRUE for even numbers.
ISFORMULA|I|reference|TRUE when the cell holds a formula.
ISLOGICAL|I|value|TRUE for logical values.
ISNA|I|value|TRUE for #N/A.
ISNONTEXT|I|value|TRUE for anything that is not text.
ISNUMBER|I|value|TRUE for numbers.
ISO.CEILING|M|number;[significance]|Rounds a number up to a multiple, whatever its sign.
ISODD|I|number|TRUE for odd numbers.
ISOWEEKNUM|D|date|ISO week number.
ISPMT|F|rate;per;nper;pv|Interest paid in a period of a loan with even principal payments.
ISREF|I|value|TRUE for references.
ISTEXT|I|value|TRUE for text.
KURT|S|number1;[number2];…|Kurtosis.
LAMBDA|L|[parameter1];…;calculation|Creates a reusable function.
LARGE|S|array;k|k-th largest value.
LCM|M|number1;[number2];…|Least common multiple.
LEFT|T|text;[num_chars]|First characters of a text.
LEFTB|T|text;[num_bytes]|First characters of a text, counting bytes.
LEN|T|text|Number of characters.
LENB|T|text|Length of a text in bytes.
LET|L|name1;value1;[name2;value2];…;calculation|Names intermediate results.
LINEST|S|known_ys;[known_xs];[const];[stats]|Linear regression statistics.
LN|M|number|Natural logarithm.
LOG|M|number;[base]|Logarithm in a base (10 by default).
LOG10|M|number|Base-10 logarithm.
LOGEST|S|known_y's;[known_x's];[const];[stats]|Statistics of an exponential curve fitted to the data.
LOGINV|S|probability;mean;standard_dev|Inverse of the lognormal distribution (compatibility).
LOGNORM.DIST|S|x;mean;standard_dev;cumulative|Lognormal distribution.
LOGNORM.INV|S|probability;mean;standard_dev|Inverse of the lognormal distribution.
LOGNORMDIST|S|x;mean;standard_dev|Cumulative lognormal distribution (compatibility).
LOOKUP|R|lookup_value;lookup_vector;[result_vector]|Looks up a value in a row or column.
LOWER|T|text|Converts text to lowercase.
MAKEARRAY|L|rows;columns;lambda|Builds an array by calling a LAMBDA for each row and column.
MAP|L|array1;[array2];…;lambda|Applies a LAMBDA to each value of the arrays.
MATCH|R|lookup_value;lookup_array;[match_type]|Position of a value in a range.
MAX|S|number1;[number2];…|Largest value.
MAXA|S|value1;[value2];…|Largest value, counting text and logicals.
MAXIFS|S|max_range;criteria_range1;criteria1;[criteria_range2;criteria2];…|Largest value among cells that meet conditions.
MDETERM|M|array|Matrix determinant.
MDURATION|F|settlement;maturity;coupon;yld;frequency;[basis]|Modified Macaulay duration of a security.
MEDIAN|S|number1;[number2];…|Median.
MID|T|text;start_num;num_chars|Characters from the middle of a text.
MIDB|T|text;start_num;num_bytes|Characters from the middle of a text, counting bytes.
MIN|S|number1;[number2];…|Smallest value.
MINA|S|value1;[value2];…|Smallest value, counting text and logicals.
MINIFS|S|min_range;criteria_range1;criteria1;[criteria_range2;criteria2];…|Smallest value among cells that meet conditions.
MINUTE|D|serial_number|Minute (0-59).
MINVERSE|M|array|Inverse matrix.
MIRR|F|values;finance_rate;reinvest_rate|Modified internal rate of return.
MMULT|M|array1;array2|Matrix product.
MOD|M|number;divisor|Remainder of a division.
MODE|S|number1;[number2];…|Most frequent value.
MODE.MULT|S|number1;[number2];…|All most-frequent values.
MODE.SNGL|S|number1;[number2];…|Most frequent value.
MONTH|D|serial_number|Month (1-12).
MROUND|M|number;multiple|Rounds to the nearest multiple.
MULTINOMIAL|M|number1;[number2];…|Multinomial of a set of numbers.
MUNIT|M|dimension|Identity matrix of a size.
N|I|value|Value converted to a number.
NA|I||The #N/A error.
NEGBINOM.DIST|S|number_f;number_s;probability_s;cumulative|Negative binomial distribution.
NEGBINOMDIST|S|number_f;number_s;probability_s|Negative binomial distribution (compatibility).
NETWORKDAYS|D|start_date;end_date;[holidays]|Working days between two dates.
NETWORKDAYS.INTL|D|start_date;end_date;[weekend];[holidays]|Working days with custom weekends.
NOMINAL|F|effect_rate;npery|Nominal annual interest rate.
NORM.DIST|S|x;mean;standard_dev;cumulative|Normal distribution.
NORM.INV|S|probability;mean;standard_dev|Inverse normal distribution.
NORM.S.DIST|S|z;cumulative|Standard normal distribution.
NORM.S.INV|S|probability|Inverse standard normal distribution.
NORMDIST|S|x;mean;standard_dev;cumulative|Normal distribution (compatibility).
NORMINV|S|probability;mean;standard_dev|Inverse of the normal distribution (compatibility).
NORMSDIST|S|z|Standard normal cumulative distribution (compatibility).
NORMSINV|S|probability|Inverse of the standard normal distribution (compatibility).
NOT|L|logical|Reverses a logical value.
NOW|D||Current date and time.
NPER|F|rate;pmt;pv;[fv];[type]|Number of payment periods.
NPV|F|rate;value1;[value2];…|Net present value.
NUMBERVALUE|T|text;[decimal_separator];[group_separator]|Converts text to a number, locale-independently.
OCT2BIN|E|number;[places]|Converts an octal number to binary.
OCT2DEC|E|number|Converts an octal number to decimal.
OCT2HEX|E|number;[places]|Converts an octal number to hexadecimal.
ODD|M|number|Rounds up to the nearest odd integer.
ODDFPRICE|F|settlement;maturity;issue;first_coupon;rate;yld;redemption;frequency;[basis]|Price per $100 face value of a security with an odd first period.
ODDFYIELD|F|settlement;maturity;issue;first_coupon;rate;pr;redemption;frequency;[basis]|Yield of a security with an odd first period.
ODDLPRICE|F|settlement;maturity;last_interest;rate;yld;redemption;frequency;[basis]|Price per $100 face value of a security with an odd last period.
ODDLYIELD|F|settlement;maturity;last_interest;rate;pr;redemption;frequency;[basis]|Yield of a security with an odd last period.
OFFSET|R|reference;rows;cols;[height];[width]|Reference offset from a starting cell.
OR|L|logical1;[logical2];…|TRUE when any argument is TRUE.
PDURATION|F|rate;pv;fv|Number of periods an investment needs to reach a value.
PEARSON|S|array1;array2|Pearson correlation coefficient.
PERCENTILE|S|array;k|k-th percentile.
PERCENTILE.EXC|S|array;k|k-th percentile, exclusive.
PERCENTILE.INC|S|array;k|k-th percentile, inclusive.
PERCENTOF|M|data_subset;data_all|Share of a subset in the total, as a fraction.
PERCENTRANK|S|array;x;[significance]|Rank of a value as a percentage.
PERCENTRANK.EXC|S|array;x;[significance]|Rank of a value as a percentage of the data (exclusive).
PERCENTRANK.INC|S|array;x;[significance]|Rank of a value as a percentage of the data (inclusive).
PERMUT|M|number;number_chosen|Number of permutations.
PERMUTATIONA|S|number;number_chosen|Number of permutations with repetitions.
PHI|S|x|Density of the standard normal distribution.
PHONETIC|T|reference|Phonetic (furigana) characters of a text.
PI|M||The number π.
PIVOTBY|R|row_fields;col_fields;values;function;[field_headers];[row_total_depth];[row_sort_order];[col_total_depth];[col_sort_order];[filter_array];[relative_to]|Groups values by row and column fields and aggregates them.
PMT|F|rate;nper;pv;[fv];[type]|Periodic loan payment.
POISSON|S|x;mean;cumulative|Poisson distribution (compatibility).
POISSON.DIST|S|x;mean;cumulative|Poisson distribution.
POWER|M|number;power|Number raised to a power.
PPMT|F|rate;per;nper;pv;[fv];[type]|Principal part of a payment.
PRICE|F|settlement;maturity;rate;yld;redemption;frequency;[basis]|Price per $100 face value of a security that pays periodic interest.
PRICEDISC|F|settlement;maturity;discount;redemption;[basis]|Price per $100 face value of a discounted security.
PRICEMAT|F|settlement;maturity;issue;rate;yld;[basis]|Price per $100 face value of a security that pays interest at maturity.
PROB|S|x_range;prob_range;lower_limit;[upper_limit]|Probability that values in a range lie between two limits.
PRODUCT|M|number1;[number2];…|Multiplies the numbers.
PROPER|T|text|Capitalizes each word.
PV|F|rate;nper;pmt;[fv];[type]|Present value of an investment.
QUARTILE|S|array;quart|Quartile of a data set.
QUARTILE.EXC|S|array;quart|Quartile of the data (exclusive).
QUARTILE.INC|S|array;quart|Quartile of a data set, inclusive.
QUOTIENT|M|numerator;denominator|Integer part of a division.
RADIANS|M|angle|Degrees to radians.
RAND|M||Random number between 0 and 1.
RANDARRAY|M|[rows];[columns];[min];[max];[whole_number]|Array of random numbers.
RANDBETWEEN|M|bottom;top|Random integer between two numbers.
RANK|S|number;ref;[order]|Rank of a number in a list.
RANK.AVG|S|number;ref;[order]|Rank, averaging ties.
RANK.EQ|S|number;ref;[order]|Rank, ties share the top rank.
RATE|F|nper;pmt;pv;[fv];[type];[guess]|Interest rate per period.
RECEIVED|F|settlement;maturity;investment;discount;[basis]|Amount received at maturity for a fully invested security.
REDUCE|L|[initial_value];array;lambda|Reduces an array to an accumulated value with a LAMBDA.
REGEXEXTRACT|T|text;pattern;[return_mode];[case_sensitivity]|Text that matches a regular expression.
REGEXREPLACE|T|text;pattern;replacement;[occurrence];[case_sensitivity]|Replaces text that matches a regular expression.
REGEXTEST|T|text;pattern;[case_sensitivity]|TRUE when text matches a regular expression.
REPLACE|T|old_text;start_num;num_chars;new_text|Replaces characters at a position.
REPLACEB|T|old_text;start_num;num_bytes;new_text|Replaces part of a text, counting bytes.
REPT|T|text;number_times|Repeats text.
RIGHT|T|text;[num_chars]|Last characters of a text.
RIGHTB|T|text;[num_bytes]|Last characters of a text, counting bytes.
ROMAN|M|number;[form]|Converts a number to a Roman numeral.
ROUND|M|number;num_digits|Rounds to a number of digits.
ROUNDDOWN|M|number;num_digits|Rounds toward zero.
ROUNDUP|M|number;num_digits|Rounds away from zero.
ROW|R|[reference]|Row number of a reference.
ROWS|R|array|Number of rows in a reference or array.
RRI|F|nper;pv;fv|Equivalent interest rate for the growth of an investment.
RSQ|S|known_ys;known_xs|Square of the correlation coefficient.
SCAN|L|[initial_value];array;lambda|Applies a LAMBDA to each value and returns every intermediate value.
SEARCH|T|find_text;within_text;[start_num]|Position of text within text (not case-sensitive, wildcards).
SEARCHB|T|find_text;within_text;[start_num]|Position of a text within another (ignoring case), counting bytes.
SEC|M|number|Secant of an angle.
SECH|M|number|Hyperbolic secant.
SECOND|D|serial_number|Second (0-59).
SEQUENCE|M|rows;[columns];[start];[step]|Array of sequential numbers.
SERIESSUM|M|x;n;m;coefficients|Sum of a power series.
SHEET|I|[value]|Sheet number.
SHEETS|I|[reference]|Number of sheets.
SIGN|M|number|Sign of a number (-1, 0 or 1).
SIN|M|number|Sine of an angle in radians.
SINH|M|number|Hyperbolic sine.
SKEW|S|number1;[number2];…|Skewness.
SKEW.P|S|number1;[number2];…|Skewness of a population.
SLN|F|cost;salvage;life|Straight-line depreciation.
SLOPE|S|known_ys;known_xs|Slope of the regression line.
SMALL|S|array;k|k-th smallest value.
SORT|R|array;[sort_index];[sort_order];[by_col]|Sorts an array.
SORTBY|R|array;by_array1;[sort_order1];…|Sorts an array by other arrays.
SQRT|M|number|Square root.
SQRTPI|M|number|Square root of a number times pi.
STANDARDIZE|S|x;mean;standard_dev|Normalized value.
STDEV|S|number1;[number2];…|Sample standard deviation.
STDEV.P|S|number1;[number2];…|Population standard deviation.
STDEV.S|S|number1;[number2];…|Sample standard deviation.
STDEVA|S|value1;[value2];…|Sample standard deviation, counting text and logicals.
STDEVP|S|number1;[number2];…|Population standard deviation.
STDEVPA|S|value1;[value2];…|Population standard deviation, counting text and logical values.
STEYX|S|known_y's;known_x's|Standard error of the predicted y values of a regression.
SUBSTITUTE|T|text;old_text;new_text;[instance_num]|Replaces occurrences of text.
SUBTOTAL|M|function_num;ref1;[ref2];…|Subtotal that ignores other subtotals.
SUM|M|number1;[number2];…|Adds the numbers.
SUMIF|M|range;criteria;[sum_range]|Sum of the cells that meet a condition.
SUMIFS|M|sum_range;criteria_range1;criteria1;[criteria_range2;criteria2];…|Sum of the cells that meet several conditions.
SUMPRODUCT|M|array1;[array2];…|Sum of the products of corresponding items.
SUMSQ|M|number1;[number2];…|Sum of squares.
SUMX2MY2|M|array_x;array_y|Sum of the differences of the squares of matching values.
SUMX2PY2|M|array_x;array_y|Sum of the sums of the squares of matching values.
SUMXMY2|M|array_x;array_y|Sum of the squared differences of matching values.
SWITCH|L|expression;value1;result1;[value2;result2];…;[default]|Matches an expression against values.
SYD|F|cost;salvage;life;per|Sum-of-years' digits depreciation.
T|T|value|Text, or empty text for other values.
T.DIST|S|x;deg_freedom;cumulative|Left-tailed Student's t distribution.
T.DIST.2T|S|x;deg_freedom|Two-tailed Student's t distribution.
T.DIST.RT|S|x;deg_freedom|Right-tailed Student's t distribution.
T.INV|S|probability;deg_freedom|Left-tailed inverse of the Student's t distribution.
T.INV.2T|S|probability;deg_freedom|Two-tailed inverse of the Student's t distribution.
T.TEST|S|array1;array2;tails;type|Probability of a Student's t-test.
TAKE|R|array;rows;[columns]|Rows or columns from the start or end of an array.
TAN|M|number|Tangent of an angle in radians.
TANH|M|number|Hyperbolic tangent.
TBILLEQ|F|settlement;maturity;discount|Bond-equivalent yield of a Treasury bill.
TBILLPRICE|F|settlement;maturity;discount|Price per $100 face value of a Treasury bill.
TBILLYIELD|F|settlement;maturity;pr|Yield of a Treasury bill.
TDIST|S|x;deg_freedom;tails|Student's t distribution (compatibility).
TEXT|T|value;format_text|Formats a number as text.
TEXTAFTER|T|text;delimiter;[instance_num];[match_mode];[match_end];[if_not_found]|Text after a delimiter.
TEXTBEFORE|T|text;delimiter;[instance_num];[match_mode];[match_end];[if_not_found]|Text before a delimiter.
TEXTJOIN|T|delimiter;ignore_empty;text1;[text2];…|Joins texts with a delimiter.
TEXTSPLIT|T|text;col_delimiter;[row_delimiter];[ignore_empty];[match_mode];[pad_with]|Splits text into an array.
TIME|D|hour;minute;second|Serial number of a time.
TIMEVALUE|D|time_text|Converts time text to a serial number.
TINV|S|probability;deg_freedom|Two-tailed inverse of the Student's t distribution (compatibility).
TOCOL|R|array;[ignore];[scan_by_column]|Array as a single column.
TODAY|D||Current date.
TOROW|R|array;[ignore];[scan_by_column]|Array as a single row.
TRANSPOSE|R|array|Swaps rows and columns.
TREND|S|known_ys;[known_xs];[new_xs];[const]|Linear-trend values.
TRIM|T|text|Removes extra spaces.
TRIMMEAN|S|array;percent|Mean excluding a percentage of outliers.
TRIMRANGE|R|range;[trim_rows];[trim_cols]|The range without its empty outer rows and columns.
TRUE|L||The logical value TRUE.
TRUNC|M|number;[num_digits]|Truncates to an integer or digits.
TTEST|S|array1;array2;tails;type|Student's t-test (compatibility).
TYPE|I|value|Type of a value as a number.
UNICHAR|T|number|Unicode character for a code point.
UNICODE|T|text|Code point of the first character.
UNIQUE|R|array;[by_col];[exactly_once]|Unique rows or columns of an array.
UPPER|T|text|Converts text to uppercase.
VALUE|T|text|Converts text to a number.
VALUETOTEXT|T|value;[format]|Value as text.
VAR|S|number1;[number2];…|Sample variance.
VAR.P|S|number1;[number2];…|Population variance.
VAR.S|S|number1;[number2];…|Sample variance.
VARA|S|value1;[value2];…|Sample variance, counting text and logical values.
VARP|S|number1;[number2];…|Population variance.
VARPA|S|value1;[value2];…|Population variance, counting text and logical values.
VDB|F|cost;salvage;life;start_period;end_period;[factor];[no_switch]|Variable-declining balance depreciation.
VLOOKUP|R|lookup_value;table_array;col_index_num;[range_lookup]|Looks down the first column and returns a value from another column.
VSTACK|R|array1;[array2];…|Stacks arrays on top of each other.
WEEKDAY|D|serial_number;[return_type]|Day of the week as a number.
WEEKNUM|D|serial_number;[return_type]|Week number of the year.
WEIBULL|S|x;alpha;beta;cumulative|Weibull distribution (compatibility).
WEIBULL.DIST|S|x;alpha;beta;cumulative|Weibull distribution.
WORKDAY|D|start_date;days;[holidays]|Date a number of working days away.
WORKDAY.INTL|D|start_date;days;[weekend];[holidays]|Working-day date with custom weekends.
WRAPCOLS|R|vector;wrap_count;[pad_with]|Wraps a vector into columns.
WRAPROWS|R|vector;wrap_count;[pad_with]|Wraps a vector into rows.
XIRR|F|values;dates;[guess]|Internal rate of return for irregular cash flows.
XLOOKUP|R|lookup_value;lookup_array;return_array;[if_not_found];[match_mode];[search_mode]|Looks up a value and returns the matching item.
XMATCH|R|lookup_value;lookup_array;[match_mode];[search_mode]|Position of a value in an array.
XNPV|F|rate;values;dates|Net present value for irregular cash flows.
XOR|L|logical1;[logical2];…|TRUE when an odd number of arguments are TRUE.
YEAR|D|serial_number|Year of a date.
YEARFRAC|D|start_date;end_date;[basis]|Fraction of the year between two dates.
YIELD|F|settlement;maturity;rate;pr;redemption;frequency;[basis]|Yield of a security that pays periodic interest.
YIELDDISC|F|settlement;maturity;pr;redemption;[basis]|Annual yield of a discounted security.
YIELDMAT|F|settlement;maturity;issue;rate;pr;[basis]|Annual yield of a security that pays interest at maturity.
Z.TEST|S|array;x;[sigma]|One-tailed probability of a z-test.
ZTEST|S|array;x;[sigma]|One-tailed probability of a z-test (compatibility).
`;

function guessCategory(name: string): string {
  if (/^(IM|BESSEL|BIN2|DEC2|HEX2|OCT2|BIT|ERF|COMPLEX|DELTA|GESTEP|CONVERT)/.test(name)) return CATEGORIES.E;
  if (/^D(AVERAGE|COUNT|GET|MAX|MIN|PRODUCT|STDEV|SUM|VAR)/.test(name)) return CATEGORIES.B;
  if (/^(COUP|ACCR|PRICE|YIELD|TBILL|DISC|INTRATE|RECEIVED|DURATION|MDURATION|DOLLARDE|DOLLARFR|FVSCHEDULE|PDURATION|RRI|ISPMT)/.test(name)) return CATEGORIES.F;
  if (/^IS|^ERROR\.TYPE|^INFO|^AREAS/.test(name)) return CATEGORIES.I;
  if (/DIST|INV|TEST|^CHI|^GAMMA|^BETA|^BINOM|^LOGNORM|^NEGBINOM|^POISSON|^EXPON|^WEIBULL|^F\.|^T\.|^Z|^CONFIDENCE|^COVAR|^PEARSON|^STEYX|^DEVSQ|^PROB|^LOGEST|^FISHER|^GAUSS|^PHI|^HYPGEOM|^CRITBINOM|^PERMUTATIONA|^PERCENTRANK|^QUARTILE|^SKEW|^VARA|^VARPA|^STDEVPA/.test(name)) return CATEGORIES.S;
  if (/B$|^ASC$|^DBCS$|^PHONETIC$|^ARRAYTOTEXT$/.test(name)) return CATEGORIES.T;
  return CATEGORIES.M;
}

function genericArgs(min: number, max: number): string[] {
  const out: string[] = [];
  const shown = Math.min(Math.max(min, 1), 4);
  const variadic = max >= 30;
  const n = variadic ? shown : max;
  for (let i = 1; i <= n; i++) out.push(i <= min ? `arg${i}` : `[arg${i}]`);
  if (variadic) out.push('…');
  return out;
}

let catalog: Map<string, FunctionInfo> | null = null;

export function functionCatalog(): Map<string, FunctionInfo> {
  if (catalog) return catalog;
  catalog = new Map();
  const described = new Map<string, FunctionInfo>();
  for (const line of TABLE.split('\n')) {
    if (!line.trim()) continue;
    const [name, cat, args, description] = line.split('|');
    described.set(name, { name, category: CATEGORIES[cat] ?? cat, args: args ? args.split(';') : [], description });
  }
  const names = new Set([...registry.keys(), 'LET', 'LAMBDA']);
  for (const name of [...names].sort()) {
    const known = described.get(name);
    if (known) {
      catalog.set(name, known);
      continue;
    }
    const def = registry.get(name);
    catalog.set(name, { name, category: guessCategory(name), args: def ? genericArgs(def.min, def.max) : [], description: '' });
  }
  return catalog;
}

export function functionInfo(name: string): FunctionInfo | undefined {
  return functionCatalog().get(name.toUpperCase());
}

/** Functions whose name starts with the prefix (case-insensitive). */
export function matchFunctions(prefix: string, limit = 60): FunctionInfo[] {
  const p = prefix.toUpperCase();
  const out: FunctionInfo[] = [];
  for (const [name, info] of functionCatalog()) {
    if (name.startsWith(p)) out.push(info);
    if (out.length >= limit) break;
  }
  return out;
}

export function functionCategories(): string[] {
  return [...new Set([...functionCatalog().values()].map((f) => f.category))].sort();
}

/** "SUM(number1, [number2], …)" */
export function signature(info: FunctionInfo): string {
  return `${info.name}(${info.args.join(', ')})`;
}

/** A short call example built from the parameter names (shown in the formula suggestions). */
export function exampleFor(info: FunctionInfo): string {
  let rangeCol = 0;
  let num = 0;
  const range = () => {
    const col = String.fromCharCode(65 + (rangeCol++ % 26));
    return `${col}2:${col}6`;
  };
  const sample = (raw: string): string | null => {
    const a = raw.replace(/^\[|\]$/g, '').toLowerCase();
    if (a === '…') return null;
    if (/criteria(?!_range)/.test(a)) return '">10"';
    if (/range|array|database|ref|values|known|data|matrix|vector|lookup_vector|result_vector/.test(a)) return range();
    if (/^value_if_true$/.test(a)) return '"Yes"';
    if (/^value_if_false|value_if_error|value_if_na|if_not_found$/.test(a)) return '"No"';
    if (/logical|condition|test/.test(a)) return 'A2>10';
    if (/find_text|old_text|delimiter|separator/.test(a)) return '"a"';
    if (/new_text|text\d*$|^text$|string/.test(a)) return '"text"';
    if (/start_date|^date$|serial_number|settlement/.test(a)) return 'A2';
    if (/end_date|maturity/.test(a)) return 'B2';
    if (/lookup_value|within_text/.test(a)) return 'A2';
    if (/col_index|row_index|index_num|column_num|row_num|^k$|nth/.test(a)) return '2';
    if (/range_lookup|exact|match_type/.test(a)) return '0';
    if (/digits|places|decimals|significance/.test(a)) return '2';
    if (/rate|probability|alpha/.test(a)) return '0.05';
    if (/nper|periods|months|days|years|count|num_chars|start_num/.test(a)) return '12';
    if (/pv|fv|pmt|principal|cost/.test(a)) return '1000';
    if (/number|value|x$|y$|base|power/.test(a)) return String([5, 10, 15, 20][num++ % 4]);
    return 'A2';
  };
  const required = info.args.filter((a) => !a.startsWith('[') && a !== '…');
  const args = required.length ? required : info.args.filter((a) => a !== '…').slice(0, 1);
  const variadic = info.args.includes('…');
  const parts = args.map(sample).filter((x): x is string => x !== null);
  if (variadic && parts.length === 1 && /^\d+$/.test(parts[0])) parts.push(String([5, 10, 15, 20][num++ % 4]), String([5, 10, 15, 20][num++ % 4]));
  return `${info.name}(${parts.join(', ')})`;
}
