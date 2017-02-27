/**
 * GS:JSON: Google Apps Script library that provides custom Google Spreadsheet formulas for work with JSON.
 * https://github.com/legionwfz/gs-json
 *
 * Copyright (C) 2017 Alexander Pereverzev
 * Released under the MIT license.
 */

/**
 * Alias for function JSARRFLATTEN.
 */
function JSARR(sources) {
  return JSARRFLATTEN.apply(this, Array.prototype.slice.call(arguments));
}

/**
 * Creates rectangular grid with one column each value of is array with shallow copy of
 * all elements in corresponding rows of sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * Sources rectangular grids could contain different count of the row
 * in this case returned rectangular grid will be created with number of
 * rows equal to maximum number of rows in source rectangular grids.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 * In case of the source is not an array, it will be copied to each row of
 * created rectangular grid.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this element will be skipped.
 *
 * JSARRCONCAT([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing additional elements to copy.
 *   @returns {Array[][]}         An rectangular grid of arrays that received the new elements.
 *
 * Logger.log(JSARRCONCAT(
 *   [[111,112,[113.1,113.2]],[121,"#JSSKIP",[123.1,123.2]]],
 *   [[211,212,213]],
 *   [311,321],
 *   410,
 *   [[true,"William White",{"phone":2316149892}],[false,"Troy White",null]]
 * ));
 *
 * [
 *   [ [111,112,[113.1,113.2],211,212,213,311,410,true ,"William White",{"phone":2316149892}] ],
 *   [ [121,    [123.2,123.2],           ,321,410,false,"Troy White"   ,null                ] ]
 * ]
 */
function JSARRCONCAT(sources) {
  return rangeReturn_(stringifyRange_(createRangeArraysConcat_.apply(this, Array.prototype.slice.call(arguments))));
}

/**
 * Creates rectangular grid with one column each value of is array with shallow copy of
 * all elements and nested arrays in corresponding rows of sources rectangular grids.
 * Any nested objects will be copied by reference, not duplicated.
 * Function recurs into nested array arguments.
 *
 * Sources rectangular grids could contain different count of the row
 * in this case returned rectangular grid will be created with number of
 * rows equal to maximum number of rows in source rectangular grids.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 * In case of the source is not an array, it will be copied to each row of
 * created rectangular grid.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * JSARRFLATTEN([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing additional elements to copy.
 *   @returns {Array[][]}         An rectangular grid of arrays that received the new elements.
 *
 * Logger.log(JSARRFLATTEN(
 *   [[111,112,[113.1,113.2]],[121,"#JSSKIP",[123.1,123.2]]],
 *   [[211,212,213]],
 *   [311,321],
 *   410,
 *   [[true,"William White",{"phone":2316149892}],[false,"Troy White",null]]
 * ));
 *
 * [
 *   [ [111,112,113.1,113.2,211,212,213,311,410,true ,"William White",{"phone":2316149892}] ],
 *   [ [121,    123.2,123.2,           ,321,410,false,"Troy White"   ,null                ] ]
 * ]
 */
function JSARRFLATTEN(sources) {
  return rangeReturn_(stringifyRange_(createRangeArraysFlatten_.apply(this, Array.prototype.slice.call(arguments))));
}

/**
 * Alias for function JSOBJPERSIST.
 */
function JSOBJ(sources) {
  return JSOBJPERSIST.apply(this, Array.prototype.slice.call(arguments));
}

/**
 * Creates rectangular grid with one column each value of is object with properties
 * defined in first row of sources rectangular grids and values as shallow copy of
 * all elements in corresponding rows of sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * Sources rectangular grids could contain different count of the row
 * in this case returned rectangular grid will be created with number of
 * rows equal to maximum number of rows in source rectangular grids.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 * In case of the source is object, it will extend object in each row of
 * created rectangular grid.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * JSOBJPERSIST([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing first row as properties names and rest rows as properties values.
 *   @returns {Array[][]}         An rectangular grid of objects that received the new properties.
 *
 * Logger.log(JSOBJPERSIST(
 *   [["firstName","middleName","lastName"],["William","Bartlett","Gallegos"],["Troy","#JSSKIP","White"]],
 *   [["birthday"],["1978-07-15"]],
 *   ["phone",2316149892,7132357707],
 *   {"client":true},
 *   [["traking"],[{"fedex":27507120}],[{"ups":"9953096047","fedex":27429384}]]
 * ));
 *
 * [
 *   [ {"firstName":"William","middleName":"Bartlett","lastName":"Gallegos","birthday":"1978-07-15","phone":2316149892,"client":true,"traking":{                   "fedex":27507120}} ],
 *   [ {"firstName":"Troy"   ,                        "lastName":"White"   ,                        "phone":7132357707,"client":true,"traking":{"ups":"9953096047","fedex":27429384}} ]
 * ]
 */
function JSOBJPERSIST(sources) {
  return rangeReturn_(stringifyRange_(createRangeObjectsPersistent_.apply(this, Array.prototype.slice.call(arguments))));
}

/**
 * Creates rectangular grid with one column each value of is an object
 * with properties defined in odds elements and values as shallow copy of
 * even elements of corresponding rows in sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * Sources rectangular grids could contain different count of the row
 * in this case returned rectangular grid will be created with number of
 * rows equal to maximum number of rows in source rectangular grids.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * JSOBJVOLAT([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing first row as properties names and rest rows as properties values.
 *   @returns {Array[][]}         An rectangular grid of objects that received the new properties.
 *
 * Logger.log(JSOBJVOLAT(
 *   [["Beetle",111,"Golf",112]],
 *   [["Jeta",211,"Passat",212,"CC",213],["Focus",221,"C-MAX","#JSSKIP","Mustang",223]],
 *   [["Tiguan",311,"Touareg",312],["Kuga",321,"Explorer",322,"Expedition",323]]
 * ));
 *
 * [
 *   [ {"Beetle":111,"Golf":112,"Jeta" :211,"Passat":212,"CC"     :213,"Tiguan":311,"Touareg" :312                 } ],
 *   [ {                        "Focus":221,            ,"Mustang":223,"Kuga"  :321,"Explorer":322,"Expedition":323} ]
 * ]
 */
function JSOBJVOLAT(sources) {
  return rangeReturn_(stringifyRange_(createRangeObjectsVolatile_.apply(this, Array.prototype.slice.call(arguments))));
}


/**
 * Creates rectangular grid with count of column equal to sources count and
 * each value of is an array with shallow copy of all elements in corresponding
 * rows of corresponding sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * JSMULTIARRCONCAT([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing additional elements to copy.
 *   @returns {Array[][]}         An rectangular grid of arrays that received the new elements.
 *
 * Logger.log(JSMULTIARRCONCAT( [[111,112,[113.1,113.2]],[121,"#JSSKIP",[123.1,123.2]]], [[211,212,213]], [311,321], [[411],[421]] ));
 *
 * [
 *   [ [111,112,[113.1,113.2]], [211,212,213], [311], [411] ],
 *   [ [121,    [123.2,123.2]], undefined    , [321], [422] ]
 * ]
 */
function JSMULTIARRCONCAT(sources) {
  return rangeReturn_(stringifyRange_(createMultiArraysConcat_.apply(this, Array.prototype.slice.call(arguments))));
}

/**
 * Creates rectangular grid with count of column equal to sources count and
 * each value of is an array with shallow copy of all elements and nested arrays
 * in corresponding rows of corresponding sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function recurs into nested array arguments.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * JSMULTIARRFLATTEN([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing additional elements to copy.
 *   @returns {Array[][]}         An rectangular grid of arrays that received the new elements.
 *
 * Logger.log(JSMULTIARRFLATTEN( [[111,112,[113.1,113.2]],[121,"#JSSKIP",[123.1,123.2]]], [[211,212,213]], [311,321], [[411],[421]] ));
 *
 * [
 *   [ [111,112,113.1,113.2], [211,212,213], [311], [411] ],
 *   [ [121,    123.2,123.2], undefined    , [321], [422] ]
 * ]
 */
function JSMULTIARRFLATTEN(sources) {
  return rangeReturn_(stringifyRange_(createMultiArraysFlatten_.apply(this, Array.prototype.slice.call(arguments))));
}

/**
 * Creates rectangular grid with count of column equal to sources count and
 * each value of is an object with properties defined in first rows of
 * sources rectangular grids and values as shallow copy of
 * all elements in corresponding rows of sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * JSMULTIOBJPERSIST([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing first row as properties names and rest rows as properties values.
 *   @returns {Array[][]}         An rectangular grid of objects that received the new properties.
 *
 * Logger.log(JSMULTIOBJPERSIST(
 *   [["firstName","middleName","lastName"],["William","Bartlett","Gallegos"],["Troy","#JSSKIP","White"]],
 *   [["birthday"],["1978-07-15"]],
 *   ["phone",2316149892,7132357707],
 *   [["traking"],[{"fedex":27507120}],[{"ups":"9953096047","fedex":27429384}]]
 * ));
 *
 * [
 *   [ {"firstName":"William","middleName":"Bartlett","lastName":"Gallegos"}, {"birthday":"1978-07-15"}, {"phone":2316149892}, {"traking":{                   "fedex":27507120}} ],
 *   [ {"firstName":"Troy"   ,                        "lastName":"White"   }, undefined                , {"phone":7132357707}, {"traking":{"ups":"9953096047","fedex":27429384}} ]
 * ]
 */
function JSMULTIOBJPERSIST(sources) {
  return rangeReturn_(stringifyRange_(createMultiObjectsPersistent_.apply(this, Array.prototype.slice.call(arguments))));
}

/**
 * Creates rectangular grid with count of column equal to sources count and
 * each value of is an object with properties defined in odds elements and
 * values as shallow copy of even elements of corresponding rows in
 * sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * JSMULTIOBJVOLAT([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing first row as properties names and rest rows as properties values.
 *   @returns {Array[][]}         An rectangular grid of objects that received the new properties.
 *
 * Logger.log(JSMULTIOBJVOLAT(
 *   [["Beetle",111,"Golf",112]],
 *   [["Jeta",211,"Passat",212,"CC",213],["Focus",221,"C-MAX","#JSSKIP","Mustang",223]],
 *   [["Tiguan",311,"Touareg",312],["Kuga",321,"Explorer",322,"Expedition",323]]
 * ));
 *
 * [
 *   [ {"Beetle":111,"Golf":112}, {"Jeta" :211,"Passat":212,"CC"     :213}, {"Tiguan":311,"Touareg" :312                 } ],
 *   [ undefined,                 {"Focus":221,            ,"Mustang":223}, {"Kuga"  :321,"Explorer":322,"Expedition":323} ]
 * ]
 */
function JSMULTIOBJVOLAT(sources) {
  return rangeReturn_(stringifyRange_(createMultiObjectsVolatile_.apply(this, Array.prototype.slice.call(arguments))));
}


/**
 * Fills rectangular grid of values with specific value or values of another rectangular grid of same size
 * depending on the presence of cells values among the array of conditional values.
 *
 * Note: If condition is omitted rectangular grid will be filled with specific value without checking any conditions.
 *
 * JSRANGEFILL(range, values [, conditions])
 *   @param {Object[][]} range       Rectangular grid of values depending on conditional values.
 *   @param {Object[][]} values      Specific value used for filling or rectangular grid of values to be used to fill corresponding cells.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Object[][]}           Filed rectangular grid depending on conditional values.
 *
 * JSRANGEFILL(arr, values [, conditions])
 *   @param {Object[]} arr           Array of values for fill depending on conditional values.
 *   @param {Object[]} values        Specific value used for filling or array of values to be used to fill corresponding elements.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Object[]}             Filed array depending on conditional values.
 *
 * JSRANGEFILL(value, values [, conditions])
 *   @param {Object} value           Value for fill depending on the presence among the conditional values.
 *   @param {Object} values          Specific value used for filling.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Object}               Filed value depending on conditional values.
 */
function JSRANGEFILL(source, values, conditions) {
  return rangeReturn_(rangeFill_.apply(this, Array.prototype.slice.call(arguments)));
}

/**
 * Returns a filtered version of the source rectangular grid for each row in conditions array,
 * returning only rows which correspond to the rows of checked values array rows of which meet the specified conditions.
 *
 * JSRANGEFILTER(range, values, conditions)
 *   @param {Object[][]} range         Rectangular grid of values depending on conditional values.
 *   @param {Object[][]} values        Array specific values used for checking on conditional values.
 *   @param {Object[][]} conditions    Array of conditional values or arrays of conditional values.
 *   @returns {Object[][]}             Filtered versions of the source rectangular grid.
 */
function JSRANGEFILTER(source, values, conditions) {
  if (arguments.length = 3) {
    throw "Wrong number of arguments to JSRANGEFILTER. Expected 3 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return rangeReturn_(stringifyRange_(rangeFilter_.apply(this, Array.prototype.slice.call(arguments))));
}

/**
 * Fills rectangular grid of values with TRUE or FALSE depending on the presence of cells values among the array of conditional values.
 *
 * JSRANGEIS(range, conditions)
 *   @param {Object[][]} range       Rectangular grid of values for fill with TRUE or FALSE depending on conditional values.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Boolean[][]}          Filed rectangular grid with TRUE or FALSE depending on conditional values.
 *
 * JSRANGEIS(arr, conditions)
 *   @param {Object[]} arr           Array of values for fill with TRUE or FALSE depending on conditional values.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Boolean[]}            Filed array with TRUE or FALSE depending on conditional values.
 *
 * JSRANGEIS(value, conditions)
 *   @param {Object} value           Value for checking the presence among the conditional values.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Boolean}              TRUE or FALSE depending on conditional values.
 */
function JSRANGEIS(source, conditions) {
  if (arguments.length !== 2) {
    throw "Wrong number of arguments to JSRANGEIS. Expected 2 arguments, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return rangeReturn_(rangeIs_.apply(this, Array.prototype.slice.call(arguments)));
}

/**
 * Trims the rectangular grid removing all blank-skip rows beginning from the bottom and all blank-skip columns beginning for the right.
 * Where blank-skip row/columns is a row/columns containing only following values: undefined, "" (empty string), "#JSSKIP" (skip-constant, see more: getSkip_()).
 *
 * JSRANGETRIM(range)
 *   @param {Object[][]} range    Rectangular grid of values to trim.
 *   @returns {Object[][]}        Trimmed rectangular grid of values.
 *
 * JSRANGETRIM(arr)
 *   @param {Object[]} arr        Array of values to trim.
 *   @returns {Object[]}          Trimmed array of values.
 *
 * JSRANGETRIM(value)
 *   @param {Object} value        Value to trim.
 *   @returns {Object}            Original value or undefined.
 */
function JSRANGETRIM(range) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to JSRANGETRIM. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return rangeReturn_(rangeTrim_.apply(this, Array.prototype.slice.call(arguments)));
}

/**
 * Fills each cell of rectangular grid with TRUE if cell value is blank value otherwise fills cell with FALSE.
 * Where blank value is a row/columns containing only following values: undefined, "" (empty string), "#JSSKIP" (skip-constant, see more: getSkip_()).
 *
 * JSISBLANK(range)
 *   @param {Object[][]} range       Rectangular grid of values for fill.
 *   @returns {Boolean[][]}          Filed rectangular grid with TRUE or FALSE.
 *
 * JSISBLANK(arr)
 *   @param {Object[]} arr           Array of values for fill.
 *   @returns {Boolean[]}            Filed array with TRUE or FALSE.
 *
 * JSISBLANK(value)
 *   @param {Object} value           Value for checking for an blank value.
 *   @returns {Boolean}              TRUE if value is blank value otherwise FALSE.
 */
function JSISBLANK(range) {
  if (arguments.length !== 1) {
    throw "Wrong number of arguments to JSISBLANK. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  var parameters = Array.prototype.slice.call(arguments);
  parameters.push([undefined, ""]);
  return rangeIs_.apply(this, parameters);
}

/**
 * Fills each cell of rectangular grid with TRUE if cell value is "#JSSKIP" (skip-constant) otherwise fills cell with FALSE.
 *
 * JSISSKIP(range)
 *   @param {Object[][]} range       Rectangular grid of values for fill.
 *   @returns {Boolean[][]}          Filed rectangular grid with TRUE or FALSE.
 *
 * JSISSKIP(arr)
 *   @param {Object[]} arr           Array of values for fill.
 *   @returns {Boolean[]}            Filed array with TRUE or FALSE.
 *
 * JSISSKIP(value)
 *   @param {Object} value           Value for checking for an blank value.
 *   @returns {Boolean}              TRUE if value is "#JSSKIP" (skip-constant) otherwise FALSE.
 */
function JSISSKIP(range) {
  if (arguments.length !== 1) {
    throw "Wrong number of arguments to JSISSKIP. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  var parameters = Array.prototype.slice.call(arguments);
  parameters.push(getSkip_());
  return rangeIs_.apply(this, parameters);
}

/**
 * Fills rectangular grid of values with "#JSSKIP" (skip-constant).
 *
 * See more getSkip_().
 *
 * JSSKIP(range)
 *   @param {Object[][]} range       Rectangular grid of values to fill.
 *   @returns {Object[][]}           Rectangular grid filed with "#JSSKIP" (skip-constant).
 *
 * JSSKIP(arr)
 *   @param {Object[]} arr           Array of values for fill.
 *   @returns {Object[]}             Array filed with "#JSSKIP" (skip-constant).
 *
 * JSSKIP(value)
 *   @param {Object} value           Any value.
 *   @returns {Object}               "#JSSKIP" (skip-constant).
 *
 * JSSKIP()
 *   @returns {Object}               "#JSSKIP" (skip-constant).
 */
function JSSKIP(range) {
  var parameters = Array.prototype.slice.call(arguments);
  if (parameters.length === 0) {
    return getSkip_();
  } else {
    parameters.push(getSkip_());
    return rangeFill_.apply(this, parameters);
  }
}

/**
 * Sleeps for 1 second and return parameter without any changes.
 * Function allows to wrap any other function and before returning result put the script to sleep.
 *
 * See more sleep_().
 *
 * JSSLEEP(value)
 *   @param {Object} value    Any value.
 *   @returns {Object}        Parameter value without any changes.
 */
function JSSLEEP(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to JSSLEEP. Expected between 0 and 1 arguments, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  sleep_(1000);

  return value;
}


/************************************  PRIVATE SECTION  *************************************/


/**
 * Defines value for representing skip-constant.
 * Skip-constant - is special value that used by different functions as indicator of necessity to skip-constant or action.
 *
 * By default skip-constant is set to "#JSSKIP" string but it could be redefined to the custom value.
 *
 * WARNING: Setting skip-constant to an object can lead to undesirable consequences since because custom Google Spreadsheet formulas can return objects.
 */
skip = "#JSSKIP";

/**
 * Copy shallowly all of the elements in the source array over to the destination array.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this element will be skipped.
 *
 * WARNING: Unlike Array.prototype.concat() method does not create new array.
 *
 * concatArray_(target, source)
 *   @param {Array} target    An array that will receive the new elements.
 *   @param {Array} source    An array containing additional elements to copy.
 *   @returns {Array}         An array that received the new elements.
 */
function concatArray_(target, source) {
  if (arguments.length < 2) {
    throw "Wrong number of arguments to CONCATARRAY_. Expected at least 2 arguments, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  if (isArray_(source)) {
    for (var i = 0; i < source.length; i++) {
      if (!isSkip_(source[i])) {
        target.push(source[i]);
      }
    }
  } else {
    if (!isSkip_(source)) {
      target.push(source);
    }
  }

  return target;
}

/**
 * Creates rectangular grid with one column each value of is array with shallow copy of
 * all elements in corresponding rows of sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * Sources rectangular grids could contain different count of the row
 * in this case returned rectangular grid will be created with number of
 * rows equal to maximum number of rows in source rectangular grids.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 * In case of the source is not an array, it will be copied to each row of
 * created rectangular grid.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this element will be skipped.
 *
 * createArraysConcat_([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing additional elements to copy.
 *   @returns {Array[][]}         An rectangular grid of arrays that received the new elements.
 *
 * Logger.log(createArraysConcat_(
 *   [[111,112,[113.1,113.2]],[121,"#JSSKIP",[123.1,123.2]]],
 *   [[211,212,213]],
 *   [311,321],
 *   410,
 *   [[true,"William White",{"phone":2316149892}],[false,"Troy White",null]]
 * ));
 *
 * [
 *   [ [111,112,[113.1,113.2],211,212,213,311,410,true ,"William White",{"phone":2316149892}] ],
 *   [ [121,    [123.2,123.2],           ,321,410,false,"Troy White"   ,null                ] ]
 * ]
 */
function createRangeArraysConcat_(sources) {
  var range = [];

  var rows = 0;
  var lengths = [];
  for (var i = 0; i < arguments.length; i++) {
    var source = parseRange_(arguments[i]);
    var length = isArray_(source) ? source.length : NaN;
    arguments[i] = source;
    lengths.push(length);
    if (!isNaN(length)) {
      rows = Math.max(length, rows);
    }
  }

  for (var j = 0; j < rows; j++) {
    range.push([
      []
    ]);
  }

  for (var i = 0; i < arguments.length; i++) {
    var source = arguments[i];
    var length = lengths[i];
    if (isNaN(length)) {
      for (var j = rows - 1; j >= 0; j--) {
        range[j][0].push(source);
      }
    } else {
      for (var j = length - 1; j >= 0; j--) {
        concatArray_(range[j][0], source[j]);
      }
    }
  }

  return range;
}

/**
 * Creates rectangular grid with one column each value of is array with shallow copy of
 * all elements and nested arrays in corresponding rows of sources rectangular grids.
 * Any nested objects will be copied by reference, not duplicated.
 * Function recurs into nested array arguments.
 *
 * Sources rectangular grids could contain different count of the row
 * in this case returned rectangular grid will be created with number of
 * rows equal to maximum number of rows in source rectangular grids.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 * In case of the source is not an array, it will be copied to each row of
 * created rectangular grid.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * createArraysFlatten_([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing additional elements to copy.
 *   @returns {Array[][]}         An rectangular grid of arrays that received the new elements.
 *
 * Logger.log(createArraysConcat_(
 *   [[111,112,[113.1,113.2]],[121,"#JSSKIP",[123.1,123.2]]],
 *   [[211,212,213]],
 *   [311,321],
 *   410,
 *   [[true,"William White",{"phone":2316149892}],[false,"Troy White",null]]
 * ));
 *
 * [
 *   [ [111,112,113.1,113.2,211,212,213,311,410,true ,"William White",{"phone":2316149892}] ],
 *   [ [121,    123.2,123.2,           ,321,410,false,"Troy White"   ,null                ] ]
 * ]
 */
function createRangeArraysFlatten_(sources) {
  var range = [];

  var rows = 0;
  var lengths = [];
  for (var i = 0; i < arguments.length; i++) {
    var source = parseRange_(arguments[i]);
    var length = isArray_(source) ? source.length : NaN;
    arguments[i] = source;
    lengths.push(length);
    if (!isNaN(length)) {
      rows = Math.max(length, rows);
    }
  }

  for (var j = 0; j < rows; j++) {
    range.push([
      []
    ]);
  }

  for (var i = 0; i < arguments.length; i++) {
    var source = arguments[i];
    var length = lengths[i];
    if (isNaN(length)) {
      for (var j = rows - 1; j >= 0; j--) {
        range[j][0].push(source);
      }
    } else {
      for (var j = length - 1; j >= 0; j--) {
        flattenArray_(range[j][0], source[j]);
      }
    }
  }

  return range;
}

/**
 * Creates rectangular grid with one column each value of is object with properties
 * defined in first row of sources rectangular grids and values as shallow copy of
 * all elements in corresponding rows of sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * Sources rectangular grids could contain different count of the row
 * in this case returned rectangular grid will be created with number of
 * rows equal to maximum number of rows in source rectangular grids.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 * In case of the source is object, it will extend object in each row of
 * created rectangular grid.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * createObjectsPersistent_([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing first row as properties names and rest rows as properties values.
 *   @returns {Array[][]}         An rectangular grid of objects that received the new properties.
 *
 * Logger.log(createObjectsPersistent_(
 *   [["firstName","middleName","lastName"],["William","Bartlett","Gallegos"],["Troy","#JSSKIP","White"]],
 *   [["birthday"],["1978-07-15"]],
 *   ["phone",2316149892,7132357707],
 *   {"client":true},
 *   [["traking"],[{"fedex":27507120}],[{"ups":"9953096047","fedex":27429384}]]
 * ));
 *
 * [
 *   [ {"firstName":"William","middleName":"Bartlett","lastName":"Gallegos","birthday":"1978-07-15","phone":2316149892,"client":true,"traking":{                   "fedex":27507120}} ],
 *   [ {"firstName":"Troy"   ,                        "lastName":"White"   ,                        "phone":7132357707,"client":true,"traking":{"ups":"9953096047","fedex":27429384}} ]
 * ]
 */
function createRangeObjectsPersistent_(sources) {
  var range = [];

  var rows = 0;
  var dimensions = [];
  for (var i = 0; i < arguments.length; i++) {
    var source = parseRange_(arguments[i]);
    var dimension = getDimension_(source);
    if (!isNaN(dimension.rows)) {
      if (dimension.rows < 2) {
        throw "Expected key-value range with at least 2 rows in parameter #" + (i + 1);
      }
    } else {
      if (!isObject_(source)) {
        throw "Expected JavaScript object in parameter #" + (i + 1);
      }
    }
    arguments[i] = source;
    dimensions.push(dimension);
    if (!isNaN(dimension.rows)) {
      rows = Math.max(dimension.rows, rows);
    }
  }

  for (var j = rows - 1; j >= 1; j--) {
    range.push({});
  }

  for (var i = 0; i < arguments.length; i++) {
    var source = arguments[i];
    var dimension = dimensions[i];
    if (isNaN(dimension.rows)) {
      for (var j = rows - 1; j >= 1; j--) {
        extendObject_(range[j - 1], source);
      }
    } else if (isNaN(dimension.columns)) {
      for (var j = dimension.rows - 1; j >= 1; j--) {
        var obj = range[j - 1];
        var key = source[0];
        var value = source[j];
        if (!isSkip_(value)) {
          obj[key] = value;
        }
      }
    } else {
      for (var j = dimension.rows - 1; j >= 1; j--) {
        var obj = range[j - 1];
        for (var k = 0; k < dimension.columns; k++) {
          var key = source[0][k];
          var value = source[j][k];
          if (!isSkip_(value)) {
            obj[key] = value;
          }
        }
      }
    }
  }

  return range;
}

/**
 * Creates rectangular grid with one column each value of is an object
 * with properties defined in odds elements and values as shallow copy of
 * even elements of corresponding rows in sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * Sources rectangular grids could contain different count of the row
 * in this case returned rectangular grid will be created with number of
 * rows equal to maximum number of rows in source rectangular grids.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * createMultiObjectsVolatile_([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing first row as properties names and rest rows as properties values.
 *   @returns {Array[][]}         An rectangular grid of objects that received the new properties.
 *
 * Logger.log(createMultiObjectsVolatile_(
 *   [["Beetle",111,"Golf",112]],
 *   [["Jeta",211,"Passat",212,"CC",213],["Focus",221,"C-MAX","#JSSKIP","Mustang",223]],
 *   [["Tiguan",311,"Touareg",312],["Kuga",321,"Explorer",322,"Expedition",323]]
 * ));
 *
 * [
 *   [ {"Beetle":111,"Golf":112,"Jeta" :211,"Passat":212,"CC"     :213,"Tiguan":311,"Touareg" :312                 } ],
 *   [ {                        "Focus":221,            ,"Mustang":223,"Kuga"  :321,"Explorer":322,"Expedition":323} ]
 * ]
 */
function createRangeObjectsVolatile_(sources) {
  var range = [];

  var rows = 0;
  var dimensions = [];
  for (var i = 0; i < arguments.length; i++) {
    var source = parseRange_(arguments[i]);
    var dimension = getDimension_(source);
    if (!isNaN(dimension.rows)) {
      if (dimension.rows < 1 || !isEven_(dimension.columns)) {
        throw "Expected key-value range with at least 1 row and an even number of columns in parameter #" + (i + 1);
      }
    } else {
      if (!isObject_(source)) {
        throw "Expected JavaScript object in parameter #" + (i + 1);
      }
    }
    arguments[i] = source;
    dimensions.push(dimension);
    if (!isNaN(dimension.rows)) {
      rows = Math.max(dimension.rows, rows);
    }
  }

  for (var j = rows - 1; j >= 0; j--) {
    range.push({});
  }

  for (var i = 0; i < arguments.length; i++) {
    var source = arguments[i];
    var dimension = dimensions[i];
    if (isNaN(dimension.rows)) {
      for (var j = rows - 1; j >= 0; j--) {
        extendObject_(range[j - 0], source);
      }
    } else {
      for (var j = dimension.rows - 1; j >= 0; j--) {
        var obj = range[j - 0];
        for (var k = 0; k < dimension.columns; k += 2) {
          var key = source[j][k + 0];
          var value = source[j][k + 1];
          if (!isSkip_(value)) {
            obj[key] = value;
          }
        }
      }
    }
  }

  return range;
}

/**
 * Creates rectangular grid with count of column equal to sources count and
 * each value of is an array with shallow copy of all elements in corresponding
 * rows of corresponding sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * createArraysConcat_([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing additional elements to copy.
 *   @returns {Array[][]}         An rectangular grid of arrays that received the new elements.
 *
 * Logger.log(createArraysConcat_( [[111,112,[113.1,113.2]],[121,"#JSSKIP",[123.1,123.2]]], [[211,212,213]], [311,321], [[411],[421]] ));
 *
 * [
 *   [ [111,112,[113.1,113.2]], [211,212,213], [311], [411] ],
 *   [ [121,    [123.2,123.2]], undefined    , [321], [422] ]
 * ]
 */
function createMultiArraysConcat_(sources) {
  var range = [];

  for (var i = 0; i < arguments.length; i++) {
    var source = parseRange_(arguments[i]);
    var dimension = getDimension_(source);
    if (!isNaN(dimension.rows)) {
      while (range.length < dimension.rows) {
        range.push([]);
      }
      for (var j = range.length - 1; j >= 0; j--) {
        range[j].push(undefined);
      }
      for (var j = dimension.rows - 1; j >= 0; j--) {
        var arr = [];
        concatArray_(arr, source[j]);
        range[j][i] = arr;
      }
    } else {
      throw "Expected range in parameter #" + (i + 1);
    }
  }

  return range;
}

/**
 * Creates rectangular grid with count of column equal to sources count and
 * each value of is an array with shallow copy of all elements and nested arrays
 * in corresponding rows of corresponding sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function recurs into nested array arguments.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * createArraysConcat_([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing additional elements to copy.
 *   @returns {Array[][]}         An rectangular grid of arrays that received the new elements.
 *
 * Logger.log(createArraysConcat_( [[111,112,[113.1,113.2]],[121,"#JSSKIP",[123.1,123.2]]], [[211,212,213]], [311,321], [[411],[421]] ));
 *
 * [
 *   [ [111,112,113.1,113.2], [211,212,213], [311], [411] ],
 *   [ [121,    123.2,123.2], undefined    , [321], [422] ]
 * ]
 */
function createMultiArraysFlatten_(sources) {
  var range = [];

  for (var i = 0; i < arguments.length; i++) {
    var source = parseRange_(arguments[i]);
    var dimension = getDimension_(source);
    if (!isNaN(dimension.rows)) {
      while (range.length < dimension.rows) {
        range.push([]);
      }
      for (var j = range.length - 1; j >= 0; j--) {
        range[j].push(undefined);
      }
      for (var j = dimension.rows - 1; j >= 0; j--) {
        var arr = [];
        flattenArray_(arr, source[j]);
        range[j][i] = arr;
      }
    } else {
      throw "Expected range in parameter #" + (i + 1);
    }
  }

  return range;
}

/**
 * Creates rectangular grid with count of column equal to sources count and
 * each value of is an object with properties defined in first rows of
 * sources rectangular grids and values as shallow copy of
 * all elements in corresponding rows of sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * In case of the source is a one-dimensional array, it will be interpreted as
 * rectangular grid with one column.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * createMultiObjectsPersistent_([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing first row as properties names and rest rows as properties values.
 *   @returns {Array[][]}         An rectangular grid of objects that received the new properties.
 *
 * Logger.log(createMultiObjectsPersistent_(
 *   [["firstName","middleName","lastName"],["William","Bartlett","Gallegos"],["Troy","#JSSKIP","White"]],
 *   [["birthday"],["1978-07-15"]],
 *   ["phone",2316149892,7132357707],
 *   [["traking"],[{"fedex":27507120}],[{"ups":"9953096047","fedex":27429384}]]
 * ));
 *
 * [
 *   [ {"firstName":"William","middleName":"Bartlett","lastName":"Gallegos"}, {"birthday":"1978-07-15"}, {"phone":2316149892}, {"traking":{                   "fedex":27507120}} ],
 *   [ {"firstName":"Troy"   ,                        "lastName":"White"   }, undefined                , {"phone":7132357707}, {"traking":{"ups":"9953096047","fedex":27429384}} ]
 * ]
 */
function createMultiObjectsPersistent_(sources) {
  var range = [];

  for (var i = 0; i < arguments.length; i++) {
    var source = parseRange_(arguments[i]);
    var dimension = getDimension_(source);
    if (dimension.rows >= 2 && !isNaN(dimension.columns)) {
      while (range.length < dimension.rows - 1) {
        range.push([]);
      }
      for (var j = range.length - 1; j >= 0; j--) {
        range[j].push(undefined);
      }
      for (var j = dimension.rows - 1; j >= 1; j--) {
        var obj = {};
        for (var k = 0; k < dimension.columns; k++) {
          var key = source[0][k];
          var value = source[j][k];
          if (!isSkip_(value)) {
            obj[key] = value;
          }
        }
        range[j - 1][i] = obj;
      }
    } else {
      throw "Expected key-value range with at least 2 rows in parameter #" + (i + 1);
    }
  }

  return range;
}

/**
 * Creates rectangular grid with count of column equal to sources count and
 * each value of is an object with properties defined in odds elements and
 * values as shallow copy of even elements of corresponding rows in
 * sources rectangular grids.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Function does not recurs into nested array arguments.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this value will be skipped.
 *
 * createMultiObjectsVolatile_([source, ...])
 *   @param {Array[][]} source    An source rectangular grid containing first row as properties names and rest rows as properties values.
 *   @returns {Array[][]}         An rectangular grid of objects that received the new properties.
 *
 * Logger.log(createMultiObjectsVolatile_(
 *   [["Beetle",111,"Golf",112]],
 *   [["Jeta",211,"Passat",212,"CC",213],["Focus",221,"C-MAX","#JSSKIP","Mustang",223]],
 *   [["Tiguan",311,"Touareg",312],["Kuga",321,"Explorer",322,"Expedition",323]]
 * ));
 *
 * [
 *   [ {"Beetle":111,"Golf":112}, {"Jeta" :211,"Passat":212,"CC"     :213}, {"Tiguan":311,"Touareg" :312                 } ],
 *   [ undefined,                 {"Focus":221,            ,"Mustang":223}, {"Kuga"  :321,"Explorer":322,"Expedition":323} ]
 * ]
 */
function createMultiObjectsVolatile_(sources) {
  var range = [];

  for (var i = 0; i < arguments.length; i++) {
    var source = parseRange_(arguments[i]);
    var dimension = getDimension_(source);
    if (dimension.rows >= 1 && isEven_(dimension.columns)) {
      while (range.length < dimension.rows - 0) {
        range.push([]);
      }
      for (var j = range.length - 1; j >= 0; j--) {
        range[j].push(undefined);
      }
      for (var j = dimension.rows - 1; j >= 0; j--) {
        var obj = {};
        for (var k = 0; k < dimension.columns; k += 2) {
          var key = source[j][k + 0];
          var value = source[j][k + 1];
          if (!isSkip_(value)) {
            obj[key] = value;
          }
        }
        range[j - 0][i] = obj;
      }
    } else {
      throw "Expected key-value range with at least 1 row and an even number of columns in parameter #" + (i + 1);
    }
  }

  return range;
}

/**
 * Merge shallowly all of the properties in the source objects over to the destination object.
 * Any nested objects or arrays will be copied by reference, not duplicated.
 * Each next source will override properties of the same name in previous arguments.
 *
 * If value of merged property is equal to skip-constant (see more getSkip_()) this property will be skipped.
 *
 * WARNING: Method does not create new object.
 *
 * extendObject_(target, source)
 *   @param {Object} target    An object that will receive the new properties.
 *   @param {Object} source    An object containing additional properties to merge in.
 *   @returns {Object}         An object that received the new properties.
 */
function extendObject_(target, source) {
  if (arguments.length < 2) {
    throw "Wrong number of arguments to EXTENDOBJECT_. Expected at least 2 arguments, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  for (var key in source) {
    var value = source[key];
    if (!isSkip_(value)) {
      target[key] = value;
    }
  }

  return target;
}

/**
 * Copy shallowly all of the elements in the source array and nested arrays over to the destination array.
 * Any nested objects will be copied by reference, not duplicated.
 * Function recurs into nested array arguments.
 *
 * If value of handled element is equal to skip-constant (see more getSkip_()) this element will be skipped.
 *
 * WARNING: Unlike Array.prototype.concat() method does not create new array.
 *
 * concatArray_(target, source)
 *   @param {Array} target    An array that will receive the new elements.
 *   @param {Array} source    An array containing additional elements to copy.
 *   @returns {Array}         An array that received the new elements.
 */
function flattenArray_(target, source) {
  if (arguments.length < 2) {
    throw "Wrong number of arguments to FLATTENARRAY_. Expected at least 2 arguments, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  if (isArray_(source)) {
    for (var i = 0; i < source.length; i++) {
      flattenArray_(target, source[i]);
    }
  } else {
    if (!isSkip_(source)) {
      target.push(source);
    }
  }

  return target;
}

/**
 * Detects dimension of value and returns object with two properties "rows" and "columns" with count of rows and columns, respectively.
 *
 * In case of the source is a rectangular grid:
 *   - rows = count of rows;
 *   - columns = count of columns.
 * In case of the source is a one-dimensional array:
 *   - rows = array length;
 *   - columns = NaN.
 * In other cases:
 *   - rows = NaN;
 *   - columns = NaN.
 *
 * getDimension_(array)
 *   @param {Object} array    Value dimension of wich to be detected.
 *   @returns {Object}        Dimension object.
 */
function getDimension_(array) {
  var dimension = {
    rows: NaN,
    columns: NaN,
    maxColumns: -1,
    minColumns: -1,
  };

  if (!isArray_(array)) {
    dimension.minColumns = NaN;
    dimension.maxColumns = NaN;
    return dimension;
  }

  dimension.rows = array.length;

  for (var i = array.length - 1; i >= 0; i--) {
    var value = array[i];
    if (isArray_(value)) {
      dimension.maxColumns = Math.max(dimension.maxColumns, value.length);
      dimension.minColumns = (dimension.minColumns === -1) ? value.length : Math.min(dimension.minColumns, value.length);
    } else {
      dimension.minColumns = NaN;
    }
  }

  if (dimension.maxColumns === -1) {
    dimension.minColumns = NaN;
    dimension.maxColumns = NaN;
  }
  if (dimension.maxColumns === dimension.minColumns) {
    dimension.columns = dimension.maxColumns;
  }

  return dimension;
}

/**
 * Returns value defined for representing skip-constant.
 * Skip-constant - is special value that used by different functions as indicator of necessity to skip value or action.
 *
 * getSkip_()
 *   @returns {Object}   Value defined for representing skip-constant.
 */
function getSkip_() {
  if (arguments.length !== 0) {
    throw "Wrong number of arguments to GETSKIP_. Expected no argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return skip;
}

/**
 * Determine whether the argument is a JavaScript array.
 *
 * isArray_(value)
 *   @param {Object} value    Value to be checked.
 *   @returns {Boolean}       TRUE if the object is a JavaScript array otherwise FALSE.
 */
function isArray_(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to ISARRAY_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return (value && value.constructor === Array);
}

/**
 * Determine whether the argument is representing a blank value.
 * Where blank value is value equal to one of the following values: undefined, "" (empty string).
 *
 * isBlank_(value)
 *   @param {Object} value    Value to be checked.
 *   @returns {Boolean}       TRUE if the object is a blank value otherwise FALSE.
 */
function isBlank_(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to ISBLANK_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return (value === undefined || value === "");
}

/**
 * Determine whether the argument is representing an even number.
 *
 * isEven_(value)
 *   @param {Object} value    Value to be checked.
 *   @returns {Boolean}       TRUE if the object is an even number otherwise FALSE.
 */
function isEven_(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to ISEVEN_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return (!(value % 2));
}

/**
 * Determine whether the argument is representing a rectangular grid.
 * Where rectangular grid is 2D array with same length of nested arrays.
 *
 * isGrid_(value)
 *   @param {Object} value    Value to be checked.
 *   @returns {Boolean}       TRUE if the object is a rectangular grid otherwise FALSE.
 */
function isGrid_(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to ISGRID_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return (!isNaN(getDimension_(value).columns));
}

/**
 * Determine whether the argument is a JavaScript object.
 *
 * isObject_(value)
 *   @param {Object} value    Value to be checked.
 *   @returns {Boolean}       TRUE if the object is a JavaScript object otherwise FALSE.
 */
function isObject_(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to ISOBJECT_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return (value !== undefined && value !== null && typeof value === "object");
}

/**
 * Determine whether the argument is representing an odd number.
 *
 * isOdd_(value)
 *   @param {Object} value    Value to be checked.
 *   @returns {Boolean}       TRUE if the object is an odd number otherwise FALSE.
 */
function isOdd_(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to ISODD_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return (!!(value % 2));
}

/**
 * Determine whether the argument is an "#JSSKIP" (skip-constant, see more: getSkip_()).
 *
 * isSkip_(value)
 *   @param {Object} value    Value to be checked.
 *   @returns {Boolean}       TRUE if the object is a "#JSSKIP" otherwise FALSE.
 */
function isSkip_(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to ISSKIP_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return (value === getSkip_());
}

/**
 * Determine whether the argument is a JavaScript string.
 *
 * isString_(value)
 *   @param {Object} value    Value to be checked.
 *   @returns {Boolean}       TRUE if the object is a JavaScript string otherwise FALSE.
 */
function isString_(value) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to ISSTRING_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  return (value !== undefined && value !== null && value.constructor === String);
}

/**
 * Parses each string in rectangular grid, constructing the JavaScript value or object described by the string.
 *
 * Note: If cell already contains JavaScript value or object the value of the cell will not be changed and cell will be skipped.
 *
 * See more: https://developer.mozilla.org/en/docs/Web/JavaScript/Reference/Global_Objects/JSON/parse
 *
 * parseRange_(range)
 *   @param {String[][]} range    Rectangular grid of values to convert to a JSON string.
 *   @returns {Object[][]}        Rectangular grid of JSON strings representing the given values.
 *
 * parseRange_(arr)
 *   @param {String[]} arr        Array of string to parse as JSON.
 *   @returns {Object[]}          Array of objects corresponding to the given JSON texts.
 *
 * parseRange_(str)
 *   @param {String} str          String to parse as JSON.
 *   @returns {Object}            Object corresponding to the given JSON text.
 */
function parseRange_(range) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to PARSERANGE_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  if (range) {
    if (isString_(range)) {
      try {
        return (range) ? JSON.parse(range) : range;
      } catch (e) {
        // SyntaxError exception if the string to parse is not valid JSON, so the value should be returned as a string.
        return range;
      }
    } else if (isArray_(range)) {
      for (var i = range.length - 1; i >= 0; i--) {
        range[i] = parseRange_(range[i]);
      }
    }
  }

  return range;
}

/**
 * Fills rectangular grid of values with specific value or values of another rectangular grid of same size
 * depending on the presence of cells values among the array of conditional values.
 *
 * Note: If condition is omitted rectangular grid will be filled with specific value without checking any conditions.
 *
 * rangeFill_(range, values [, conditions])
 *   @param {Object[][]} range       Rectangular grid of values depending on conditional values.
 *   @param {Object[][]} values      Specific value used for filling or rectangular grid of values to be used to fill corresponding cells.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Object[][]}           Filed rectangular grid depending on conditional values.
 *
 * rangeFill_(arr, values [, conditions])
 *   @param {Object[]} arr           Array of values for fill depending on conditional values.
 *   @param {Object[]} values        Specific value used for filling or array of values to be used to fill corresponding elements.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Object[]}             Filed array depending on conditional values.
 *
 * rangeFill_(value, values [, conditions])
 *   @param {Object} value           Value for fill depending on the presence among the conditional values.
 *   @param {Object} values          Specific value used for filling.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Object}               Filed value depending on conditional values.
 */
function rangeFill_(source, values, conditions) {
  if (arguments.length < 2 || arguments.length > 3) {
    throw "Wrong number of arguments to RANGEFILL_. Expected between 2 and 3 arguments, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  var range = source;
  var setElement_ = (isArray_(conditions)) ? function(element) {
    for (var i = conditions.length - 1; i >= 0; i--) {
      if (element != conditions[i]) {
        return element;
      }
    }
    return values;
  } : function(element) {
    return (element == conditions || conditions === undefined) ? values : element;
  };

  if (isArray_(range)) {
    return range.map(function(element) {
      return (isArray_(element)) ? element.map(setElement_) : setElement_(element);
    });
  } else {
    return setElement_(range);
  }
}

/**
 * Returns a filtered version of the source rectangular grid for each row in conditions array,
 * returning only rows which correspond to the rows of checked values array rows of which meet the specified conditions.
 *
 * rangeFilter_(range, values, conditions)
 *   @param {Object[][]} range         Rectangular grid of values depending on conditional values.
 *   @param {Object[][]} values        Array specific values used for checking on conditional values.
 *   @param {Object[][]} conditions    Array of conditional values or arrays of conditional values.
 *   @returns {Object[][]}             Filtered versions of the source rectangular grid.
 */
function rangeFilter_(source, values, conditions) {
  if (arguments.length = 3) {
    throw "Wrong number of arguments to RANGEFILTER_. Expected 3 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  source = parseRange_(source);
  values = parseRange_(values);
  var sourceDimension = getDimension_(source);
  var valuesDimension = getDimension_(values);
  if (sourceDimension.rows !== valuesDimension.rows && isNaN(sourceDimension.rows) !== isNaN(valuesDimension.rows)) {
    throw "Function RANGEFILTER_ parameter 2 has mismatched row size. Expected: {0}. Actual: {1}."
    .replace("{0}", isNaN(sourceDimension.rows) ? 1 : sourceDimension.rows)
      .replace("{1}", isNaN(valuesDimension.rows) ? 1 : valuesDimension.rows);
  }

  var rows;
  var columns;
  var filters = [];
  for (var i = 2; i < arguments.length; i++) {
    var filter = parseRange_(arguments[i]);
    var dimension = getDimension_(filter);
    if (rows === undefined) {
      rows = dimension.rows;
      columns = dimension.columns;
    } else if (rows !== dimension.rows && isNaN(rows) !== isNaN(dimension.rows)) {}
    if (isNaN(rows)) {
      if (valuesDimension.columns !== 1 && isNaN(valuesDimension.columns) !== true) {
        throw "Function RANGEFILTER_ parameter 3 has mismatched column size. Expected: {0}. Actual: {1}."
        .replace("{0}", isNaN(valuesDimension.columns) ? 1 : valuesDimension.columns)
          .replace("{1}", isNaN(dimension.columns) ? 1 : dimension.columns);
      }
    } else if (isNaN(sourceDimension.rows)) {
      if (dimension.columns !== 1 && isNaN(dimension.columns) !== true) {
        throw "Function RANGEFILTER_ parameter 3 has mismatched column size. Expected: {0}. Actual: {1}."
        .replace("{0}", isNaN(valuesDimension.columns) ? 1 : valuesDimension.columns)
          .replace("{1}", isNaN(dimension.columns) ? 1 : dimension.columns);
      }
    } else {
      if (valuesDimension.minColumns !== dimension.minColumns && isNaN(valuesDimension.minColumns) !== isNaN(dimension.minColumns)) {
        throw "Function RANGEFILTER_ parameter 3 has mismatched column size. Expected: {0}. Actual: {1}."
        .replace("{0}", isNaN(valuesDimension.columns) ? 1 : valuesDimension.columns)
          .replace("{1}", isNaN(dimension.columns) ? 1 : dimension.columns);
        throw "Wrong size 31 " + JSON.stringify(valuesDimension) + " " + JSON.stringify(dimension) + " " + JSON.stringify(filter);
      }
      if (valuesDimension.maxColumns !== dimension.maxColumns && isNaN(valuesDimension.maxColumns) !== isNaN(dimension.maxColumns)) {
        throw "Function RANGEFILTER_ parameter 3 has mismatched column size. Expected: {0}. Actual: {1}."
        .replace("{0}", isNaN(valuesDimension.columns) ? 1 : valuesDimension.columns)
          .replace("{1}", isNaN(dimension.columns) ? 1 : dimension.columns);
        throw "Wrong size 31 " + JSON.stringify(valuesDimension) + " " + JSON.stringify(dimension) + " " + JSON.stringify(filter);
      }
    }
    filters.push(filter);
  }

  var range;
  if (isNaN(rows) && isNaN(sourceDimension.rows)) {
    var valid = true;
    for (var i = 0; i < filters.length && valid; i++) {
      var filter = filters[i];
      valid = (values == filter);
    }
    if (valid) {
      range = source;
    } else {
      range = undefined;
    }
  } else if (isNaN(rows)) {
    range = [];
    var filtered = [];
    range.push(filtered);
    for (var j = 0; j < sourceDimension.rows; j++) {
      var valid = true;
      for (var i = 0; i < filters.length && valid; i++) {
        var filter = filters[i];
        if (isNaN(sourceDimension.columns)) {
          valid = (values[j] == filter);
        } else {
          valid = (values[j][0] == filter);
        }
      }
      if (valid) {
        filtered.push(source[j]);
      }
    }
  } else if (isNaN(sourceDimension.rows)) {
    range = [];
    for (var l = 0; l < rows; l++) {
      var filtered = [];
      range.push(filtered);
      var valid = true;
      for (var i = 0; i < filters.length && valid; i++) {
        var filter = filters[i];
        if (isNaN(columns)) {
          valid = (values == filter[l]);
        } else {
          valid = (values == filter[l][0]);
        }
      }
      if (valid) {
        filtered.push(source);
      }
    }
  } else {
    range = [];
    for (var l = 0; l < rows; l++) {
      var filtered = [];
      range.push(filtered);
      for (var j = 0; j < sourceDimension.rows; j++) {
        var valid = true;
        for (var i = 0; i < filters.length && valid; i++) {
          var filter = filters[i];
          if (isNaN(columns)) {
            valid = (values[j] == filter[l]);
          } else {
            for (var k = 0; k < columns && valid; k++) {
              valid = (values[j][k] == filter[l][k]);
            }
          }
        }
        if (valid) {
          filtered.push(source[j]);
        }
      }
    }
  }

  return range;
}

/**
 * Fills rectangular grid of values with TRUE or FALSE depending on the presence of cells values among the array of conditional values.
 *
 * rangeIs_(range, conditions)
 *   @param {Object[][]} range       Rectangular grid of values for fill with TRUE or FALSE depending on conditional values.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Boolean[][]}          Filed rectangular grid with TRUE or FALSE depending on conditional values.
 *
 * rangeIs_(arr, conditions)
 *   @param {Object[]} arr           Array of values for fill with TRUE or FALSE depending on conditional values.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Boolean[]}            Filed array with TRUE or FALSE depending on conditional values.
 *
 * rangeIs_(value, conditions)
 *   @param {Object} value           Value for checking the presence among the conditional values.
 *   @param {Object[]} conditions    Conditional value or array of conditional values.
 *   @returns {Boolean}              TRUE or FALSE depending on conditional values.
 */
function rangeIs_(source, conditions) {
  if (arguments.length !== 2) {
    throw "Wrong number of arguments to RANGEIS_. Expected 2 arguments, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  var range = source;
  var checkElement_ = (isArray_(conditions)) ? function(element) {
    for (var i = conditions.length - 1; i >= 0; i--) {
      if (element != conditions[i]) {
        return false;
      }
    }
    return true;
  } : function(element) {
    return (element == conditions)
  };

  if (isArray_(range)) {
    return range.map(function(element) {
      return (isArray_(element)) ? element.map(checkElement_) : checkElement_(element);
    });
  } else {
    return checkElement_(range);
  }
}

/**
 * Prepares rectangular grid of values for returning as a result of custom Google Spreadsheet formula.
 * The undefined value should be returned instead of rectangular grid if grid satisfy any of the following conditions:
 *   - array (root array) is empty;
 *   - all nested array is empty.
 * Otherwise it will cause to #REF! (Reference does not exist) exception.
 *
 * rangeReturn_(range)
 *   @param {Object[][]} range    Rectangular grid of values for returning as a result of custom Google Spreadsheet formula.
 *   @returns {Object}            Original rectangular grid or undefined value depending on grid values.
 *
 * rangeReturn_(arr)
 *   @param {Object[]} arr        Array of values for returning as a result of custom Google Spreadsheet formula.
 *   @returns {Object}            Original array of values or undefined value depending on array length.
 *
 * rangeReturn_(value)
 *   @param {Object} value        Value for returning as a result of custom Google Spreadsheet formula.
 *   @returns {Object}            Original value without any changes.
 */
function rangeReturn_(range) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to RANGERERURN_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  if (!isArray_(range)) {
    return range;
  } else {
    return (!range.length) ? undefined : (!range.some(function(value) {
      return (!isArray_(value) || value.length);
    })) ? undefined : range;
  }
}

/**
 * Trims the rectangular grid removing all blank-skip rows beginning from the bottom and all blank-skip columns beginning for the right.
 * Where blank-skip row/columns is a row/columns containing only following values: undefined, "" (empty string), "#JSSKIP" (skip-constant, see more: getSkip_()).
 *
 * rangeTrim_(range)
 *   @param {Object[][]} range    Rectangular grid of values to trim.
 *   @returns {Object[][]}        Trimmed rectangular grid of values.
 *
 * rangeTrim_(arr)
 *   @param {Object[]} arr        Array of values to trim.
 *   @returns {Object[]}          Trimmed array of values.
 *
 * rangeTrim_(value)
 *   @param {Object} value        Value to trim.
 *   @returns {Object}            Original value or undefined.
 */
function rangeTrim_(range) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to RANGETRIM_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  var dimension = getDimension_(range);
  if (isNaN(dimension.rows)) {
    if (!isBlank_(value) && !isSkip_(value)) {
      return range;
    } else {
      return undefined;
    }
  }

  if (!isNaN(dimension.columns)) {
    for (var i = dimension.rows - 1; i >= 0; i--) {
      for (var j = dimension.columns - 1; j >= 0; j--) {
        var value = range[i][j];
        if (!isBlank_(value) && !isSkip_(value)) {
          return range;
        }
      }
      range.pop();
    }
  } else {
    for (var i = dimension.rows - 1; i >= 0; i--) {
      var value = range[i];
      if (!isBlank_(value) && !isSkip_(value)) {
        return range;
      }
      range.pop();
    }
  }

  return range;
}

/**
 * Converts each value in rectangular grid of a JavaScript values or objects to a JSON string.
 *
 * See more: https://developer.mozilla.org/en/docs/Web/JavaScript/Reference/Global_Objects/JSON/stringify
 *
 * stringifyRange_(range)
 *   @param {Object[][]} range    Rectangular grid of values to convert to a JSON strings.
 *   @returns {String[][]}        Rectangular grid of JSON strings representing the given values.
 *
 * stringifyRange_(arr)
 *   @param {Object[]} arr        Array of values to convert to a JSON strings.
 *   @returns {String[]}          Array of JSON strings representing the given values.
 *
 * stringifyRange_(value)
 *   @param {Object} value        Value to convert to a JSON string.
 *   @returns {String}            JSON string representing the given value.
 */
function stringifyRange_(range) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to STRINGIFYRANGE_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  if (!isArray_(range)) {
    return stringifyValue_(range);
  } else {
    return range.map(function(value) {
      return (isArray_(value)) ? value.map(stringifyValue_) : stringifyValue_(value);
    });
  }
}

/**
 * Converts a JavaScript value to a JSON string.
 *
 * See more: https://developer.mozilla.org/en/docs/Web/JavaScript/Reference/Global_Objects/JSON/stringify
 *
 * stringifyValue_(value)
 *   @param {Object} value    Value to convert to a JSON string.
 *   @returns {String}        JSON string representing the given value.
 */
function stringifyValue_(value) {
  //if (arguments.length > 1) {
  //    throw "Wrong number of arguments to STRINGIFYVALUE_. Expected 1 argument, but got {0} arguments."
  //        .replace("{0}", arguments.length);
  //}

  return JSON.stringify(value);
}

/**
 * Sleeps for specified number of milliseconds.
 * Immediately puts the script to sleep for the specified number of milliseconds.
 * The maximum allowed value is 300000 (or 5 minutes).
 *
 * Google recommends to use in custom Google Spreadsheet formulas if you expect that the formula will be used frequently.
 * See more: https://developers.google.com/apps-script/reference/utilities/utilities#sleep(Integer)
 *
 * sleep_(milliseconds)
 *   @param {Number} milliseconds    Number of milliseconds to sleep for (1000 by default).
 */
function sleep_(milliseconds) {
  if (arguments.length > 1) {
    throw "Wrong number of arguments to SLEEP_. Expected 1 argument, but got {0} arguments."
    .replace("{0}", arguments.length);
  }

  milliseconds = Math.round(milliseconds) || 1000; // By default Google recommends use 1 second pause.
  if (isNaN(milliseconds)) {
    throw "Wrong argument type to SLEEP_. Expected Number in parameter #1.";
  }

  Utilities.sleep(milliseconds);
}