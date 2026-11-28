var version = 31;
// Generated with npm run build. Edit the source files under src/.

// Source: vendor/ical.js
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/* istanbul ignore next */
/* jshint ignore:start */
if (typeof module === 'object') {
  // CommonJS, where exports may be different each time.
  ICAL = module.exports;
} else if (typeof ICAL !== 'object') {/* istanbul ignore next */
  /** @ignore */
  this.ICAL = {};
}
/* jshint ignore:end */


/**
 * The number of characters before iCalendar line folding should occur
 * @type {Number}
 * @default 75
 */
ICAL.foldLength = 75;


/**
 * The character(s) to be used for a newline. The default value is provided by
 * rfc5545.
 * @type {String}
 * @default "\r\n"
 */
ICAL.newLineChar = '\r\n';


/**
 * Helper functions used in various places within ical.js
 * @namespace
 */
ICAL.helpers = {
  /**
   * Compiles a list of all referenced TZIDs in all subcomponents and
   * removes any extra VTIMEZONE subcomponents. In addition, if any TZIDs
   * are referenced by a component, but a VTIMEZONE does not exist,
   * an attempt will be made to generate a VTIMEZONE using ICAL.TimezoneService.
   *
   * @param {ICAL.Component} vcal     The top-level VCALENDAR component.
   * @return {ICAL.Component}         The ICAL.Component that was passed in.
   */
  updateTimezones: function(vcal) {
    var allsubs, properties, vtimezones, reqTzid, i, tzid;

    if (!vcal || vcal.name !== "vcalendar") {
      //not a top-level vcalendar component
      return vcal;
    }

    //Store vtimezone subcomponents in an object reference by tzid.
    //Store properties from everything else in another array
    allsubs = vcal.getAllSubcomponents();
    properties = [];
    vtimezones = {};
    for (i = 0; i < allsubs.length; i++) {
      if (allsubs[i].name === "vtimezone") {
        tzid = allsubs[i].getFirstProperty("tzid").getFirstValue();
        vtimezones[tzid] = allsubs[i];
      } else {
        properties = properties.concat(allsubs[i].getAllProperties());
      }
    }

    //create an object with one entry for each required tz
    reqTzid = {};
    for (i = 0; i < properties.length; i++) {
      if ((tzid = properties[i].getParameter("tzid"))) {
        reqTzid[tzid] = true;
      }
    }

    //delete any vtimezones that are not on the reqTzid list.
    for (i in vtimezones) {
      if (vtimezones.hasOwnProperty(i) && !reqTzid[i]) {
        vcal.removeSubcomponent(vtimezones[i]);
      }
    }

    //create any missing, but registered timezones
    for (i in reqTzid) {
      if (
        reqTzid.hasOwnProperty(i) &&
        !vtimezones[i] &&
        ICAL.TimezoneService.has(i) &&            
        ICAL.TimezoneService.get(i).component
        
      ) {
        vcal.addSubcomponent(ICAL.TimezoneService.get(i).component);
      }
    }

    return vcal;
  },

  /**
   * Checks if the given type is of the number type and also NaN.
   *
   * @param {Number} number     The number to check
   * @return {Boolean}          True, if the number is strictly NaN
   */
  isStrictlyNaN: function(number) {
    return typeof(number) === 'number' && isNaN(number);
  },

  /**
   * Parses a string value that is expected to be an integer, when the valid is
   * not an integer throws a decoration error.
   *
   * @param {String} string     Raw string input
   * @return {Number}           Parsed integer
   */
  strictParseInt: function(string) {
    var result = parseInt(string, 10);

    if (ICAL.helpers.isStrictlyNaN(result)) {
      throw new Error(
        'Could not extract integer from "' + string + '"'
      );
    }

    return result;
  },

  /**
   * Creates or returns a class instance of a given type with the initialization
   * data if the data is not already an instance of the given type.
   *
   * @example
   * var time = new ICAL.Time(...);
   * var result = ICAL.helpers.formatClassType(time, ICAL.Time);
   *
   * (result instanceof ICAL.Time)
   * // => true
   *
   * result = ICAL.helpers.formatClassType({}, ICAL.Time);
   * (result isntanceof ICAL.Time)
   * // => true
   *
   *
   * @param {Object} data       object initialization data
   * @param {Object} type       object type (like ICAL.Time)
   * @return {?}                An instance of the found type.
   */
  formatClassType: function formatClassType(data, type) {
    if (typeof(data) === 'undefined') {
      return undefined;
    }

    if (data instanceof type) {
      return data;
    }
    return new type(data);
  },

  /**
   * Identical to indexOf but will only match values when they are not preceded
   * by a backslash character.
   *
   * @param {String} buffer         String to search
   * @param {String} search         Value to look for
   * @param {Number} pos            Start position
   * @return {Number}               The position, or -1 if not found
   */
  unescapedIndexOf: function(buffer, search, pos) {
    while ((pos = buffer.indexOf(search, pos)) !== -1) {
      if (pos > 0 && buffer[pos - 1] === '\\') {
        pos += 1;
      } else {
        return pos;
      }
    }
    return -1;
  },

  /**
   * Find the index for insertion using binary search.
   *
   * @param {Array} list            The list to search
   * @param {?} seekVal             The value to insert
   * @param {function(?,?)} cmpfunc The comparison func, that can
   *                                  compare two seekVals
   * @return {Number}               The insert position
   */
  binsearchInsert: function(list, seekVal, cmpfunc) {
    if (!list.length)
      return 0;

    var low = 0, high = list.length - 1,
        mid, cmpval;

    while (low <= high) {
      mid = low + Math.floor((high - low) / 2);
      cmpval = cmpfunc(seekVal, list[mid]);

      if (cmpval < 0)
        high = mid - 1;
      else if (cmpval > 0)
        low = mid + 1;
      else
        break;
    }

    if (cmpval < 0)
      return mid; // insertion is displacing, so use mid outright.
    else if (cmpval > 0)
      return mid + 1;
    else
      return mid;
  },

  /**
   * Convenience function for debug output
   * @private
   */
  dumpn: /* istanbul ignore next */ function() {
    if (!ICAL.debug) {
      return;
    }

    if (typeof (console) !== 'undefined' && 'log' in console) {
      ICAL.helpers.dumpn = function consoleDumpn(input) {
        console.log(input);
      };
    } else {
      ICAL.helpers.dumpn = function geckoDumpn(input) {
        dump(input + '\n');
      };
    }

    ICAL.helpers.dumpn(arguments[0]);
  },

  /**
   * Clone the passed object or primitive. By default a shallow clone will be
   * executed.
   *
   * @param {*} aSrc            The thing to clone
   * @param {Boolean=} aDeep    If true, a deep clone will be performed
   * @return {*}                The copy of the thing
   */
  clone: function(aSrc, aDeep) {
    if (!aSrc || typeof aSrc != "object") {
      return aSrc;
    } else if (aSrc instanceof Date) {
      return new Date(aSrc.getTime());
    } else if ("clone" in aSrc) {
      return aSrc.clone();
    } else if (Array.isArray(aSrc)) {
      var arr = [];
      for (var i = 0; i < aSrc.length; i++) {
        arr.push(aDeep ? ICAL.helpers.clone(aSrc[i], true) : aSrc[i]);
      }
      return arr;
    } else {
      var obj = {};
      for (var name in aSrc) {
        // uses prototype method to allow use of Object.create(null);
        /* istanbul ignore else */
        if (Object.prototype.hasOwnProperty.call(aSrc, name)) {
          if (aDeep) {
            obj[name] = ICAL.helpers.clone(aSrc[name], true);
          } else {
            obj[name] = aSrc[name];
          }
        }
      }
      return obj;
    }
  },

  /**
   * Performs iCalendar line folding. A line ending character is inserted and
   * the next line begins with a whitespace.
   *
   * @example
   * SUMMARY:This line will be fold
   *  ed right in the middle of a word.
   *
   * @param {String} aLine      The line to fold
   * @return {String}           The folded line
   */
  foldline: function foldline(aLine) {
    var result = "";
    var line = aLine || "";

    while (line.length) {
      result += ICAL.newLineChar + " " + line.substr(0, ICAL.foldLength);
      line = line.substr(ICAL.foldLength);
    }
    return result.substr(ICAL.newLineChar.length + 1);
  },

  /**
   * Pads the given string or number with zeros so it will have at least two
   * characters.
   *
   * @param {String|Number} data    The string or number to pad
   * @return {String}               The number padded as a string
   */
  pad2: function pad(data) {
    if (typeof(data) !== 'string') {
      // handle fractions.
      if (typeof(data) === 'number') {
        data = parseInt(data);
      }
      data = String(data);
    }

    var len = data.length;

    switch (len) {
      case 0:
        return '00';
      case 1:
        return '0' + data;
      default:
        return data;
    }
  },

  /**
   * Truncates the given number, correctly handling negative numbers.
   *
   * @param {Number} number     The number to truncate
   * @return {Number}           The truncated number
   */
  trunc: function trunc(number) {
    return (number < 0 ? Math.ceil(number) : Math.floor(number));
  },

  /**
   * Poor-man's cross-browser inheritance for JavaScript. Doesn't support all
   * the features, but enough for our usage.
   *
   * @param {Function} base     The base class constructor function.
   * @param {Function} child    The child class constructor function.
   * @param {Object} extra      Extends the prototype with extra properties
   *                              and methods
   */
  inherits: function(base, child, extra) {
    function F() {}
    F.prototype = base.prototype;
    child.prototype = new F();

    if (extra) {
      ICAL.helpers.extend(extra, child.prototype);
    }
  },

  /**
   * Poor-man's cross-browser object extension. Doesn't support all the
   * features, but enough for our usage. Note that the target's properties are
   * not overwritten with the source properties.
   *
   * @example
   * var child = ICAL.helpers.extend(parent, {
   *   "bar": 123
   * });
   *
   * @param {Object} source     The object to extend
   * @param {Object} target     The object to extend with
   * @return {Object}           Returns the target.
   */
  extend: function(source, target) {
    for (var key in source) {
      var descr = Object.getOwnPropertyDescriptor(source, key);
      if (descr && !Object.getOwnPropertyDescriptor(target, key)) {
        Object.defineProperty(target, key, descr);
      }
    }
    return target;
  }
};
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */

/** @namespace ICAL */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.design = (function() {
  'use strict';

  var FROM_ICAL_NEWLINE = /\\\\|\\;|\\,|\\[Nn]/g;
  var TO_ICAL_NEWLINE = /\\|;|,|\n/g;
  var FROM_VCARD_NEWLINE = /\\\\|\\,|\\[Nn]/g;
  var TO_VCARD_NEWLINE = /\\|,|\n/g;

  function createTextType(fromNewline, toNewline) {
    var result = {
      matches: /.*/,

      fromICAL: function(aValue, structuredEscape) {
        return replaceNewline(aValue, fromNewline, structuredEscape);
      },

      toICAL: function(aValue, structuredEscape) {
        var regEx = toNewline;
        if (structuredEscape)
          regEx = new RegExp(regEx.source + '|' + structuredEscape);
        return aValue.replace(regEx, function(str) {
          switch (str) {
          case "\\":
            return "\\\\";
          case ";":
            return "\\;";
          case ",":
            return "\\,";
          case "\n":
            return "\\n";
          /* istanbul ignore next */
          default:
            return str;
          }
        });
      }
    };
    return result;
  }

  // default types used multiple times
  var DEFAULT_TYPE_TEXT = { defaultType: "text" };
  var DEFAULT_TYPE_TEXT_MULTI = { defaultType: "text", multiValue: "," };
  var DEFAULT_TYPE_TEXT_STRUCTURED = { defaultType: "text", structuredValue: ";" };
  var DEFAULT_TYPE_INTEGER = { defaultType: "integer" };
  var DEFAULT_TYPE_DATETIME_DATE = { defaultType: "date-time", allowedTypes: ["date-time", "date"] };
  var DEFAULT_TYPE_DATETIME = { defaultType: "date-time" };
  var DEFAULT_TYPE_URI = { defaultType: "uri" };
  var DEFAULT_TYPE_UTCOFFSET = { defaultType: "utc-offset" };
  var DEFAULT_TYPE_RECUR = { defaultType: "recur" };
  var DEFAULT_TYPE_DATE_ANDOR_TIME = { defaultType: "date-and-or-time", allowedTypes: ["date-time", "date", "text"] };

  function replaceNewlineReplace(string) {
    switch (string) {
      case "\\\\":
        return "\\";
      case "\\;":
        return ";";
      case "\\,":
        return ",";
      case "\\n":
      case "\\N":
        return "\n";
      /* istanbul ignore next */
      default:
        return string;
    }
  }

  function replaceNewline(value, newline, structuredEscape) {
    // avoid regex when possible.
    if (value.indexOf('\\') === -1) {
      return value;
    }
    if (structuredEscape)
      newline = new RegExp(newline.source + '|\\\\' + structuredEscape);
    return value.replace(newline, replaceNewlineReplace);
  }

  var commonProperties = {
    "categories": DEFAULT_TYPE_TEXT_MULTI,
    "url": DEFAULT_TYPE_URI,
    "version": DEFAULT_TYPE_TEXT,
    "uid": DEFAULT_TYPE_TEXT
  };

  var commonValues = {
    "boolean": {
      values: ["TRUE", "FALSE"],

      fromICAL: function(aValue) {
        switch (aValue) {
          case 'TRUE':
            return true;
          case 'FALSE':
            return false;
          default:
            //TODO: parser warning
            return false;
        }
      },

      toICAL: function(aValue) {
        if (aValue) {
          return 'TRUE';
        }
        return 'FALSE';
      }

    },
    float: {
      matches: /^[+-]?\d+\.\d+$/,

      fromICAL: function(aValue) {
        var parsed = parseFloat(aValue);
        if (ICAL.helpers.isStrictlyNaN(parsed)) {
          // TODO: parser warning
          return 0.0;
        }
        return parsed;
      },

      toICAL: function(aValue) {
        return String(aValue);
      }
    },
    integer: {
      fromICAL: function(aValue) {
        var parsed = parseInt(aValue);
        if (ICAL.helpers.isStrictlyNaN(parsed)) {
          return 0;
        }
        return parsed;
      },

      toICAL: function(aValue) {
        return String(aValue);
      }
    },
    "utc-offset": {
      toICAL: function(aValue) {
        if (aValue.length < 7) {
          // no seconds
          // -0500
          return aValue.substr(0, 3) +
                 aValue.substr(4, 2);
        } else {
          // seconds
          // -050000
          return aValue.substr(0, 3) +
                 aValue.substr(4, 2) +
                 aValue.substr(7, 2);
        }
      },

      fromICAL: function(aValue) {
        if (aValue.length < 6) {
          // no seconds
          // -05:00
          return aValue.substr(0, 3) + ':' +
                 aValue.substr(3, 2);
        } else {
          // seconds
          // -05:00:00
          return aValue.substr(0, 3) + ':' +
                 aValue.substr(3, 2) + ':' +
                 aValue.substr(5, 2);
        }
      },

      decorate: function(aValue) {
        return ICAL.UtcOffset.fromString(aValue);
      },

      undecorate: function(aValue) {
        return aValue.toString();
      }
    }
  };

  var icalParams = {
    // Although the syntax is DQUOTE uri DQUOTE, I don't think we should
    // enfoce anything aside from it being a valid content line.
    //
    // At least some params require - if multi values are used - DQUOTEs
    // for each of its values - e.g. delegated-from="uri1","uri2"
    // To indicate this, I introduced the new k/v pair
    // multiValueSeparateDQuote: true
    //
    // "ALTREP": { ... },

    // CN just wants a param-value
    // "CN": { ... }

    "cutype": {
      values: ["INDIVIDUAL", "GROUP", "RESOURCE", "ROOM", "UNKNOWN"],
      allowXName: true,
      allowIanaToken: true
    },

    "delegated-from": {
      valueType: "cal-address",
      multiValue: ",",
      multiValueSeparateDQuote: true
    },
    "delegated-to": {
      valueType: "cal-address",
      multiValue: ",",
      multiValueSeparateDQuote: true
    },
    // "DIR": { ... }, // See ALTREP
    "encoding": {
      values: ["8BIT", "BASE64"]
    },
    // "FMTTYPE": { ... }, // See ALTREP
    "fbtype": {
      values: ["FREE", "BUSY", "BUSY-UNAVAILABLE", "BUSY-TENTATIVE"],
      allowXName: true,
      allowIanaToken: true
    },
    // "LANGUAGE": { ... }, // See ALTREP
    "member": {
      valueType: "cal-address",
      multiValue: ",",
      multiValueSeparateDQuote: true
    },
    "partstat": {
      // TODO These values are actually different per-component
      values: ["NEEDS-ACTION", "ACCEPTED", "DECLINED", "TENTATIVE",
               "DELEGATED", "COMPLETED", "IN-PROCESS"],
      allowXName: true,
      allowIanaToken: true
    },
    "range": {
      values: ["THISLANDFUTURE"]
    },
    "related": {
      values: ["START", "END"]
    },
    "reltype": {
      values: ["PARENT", "CHILD", "SIBLING"],
      allowXName: true,
      allowIanaToken: true
    },
    "role": {
      values: ["REQ-PARTICIPANT", "CHAIR",
               "OPT-PARTICIPANT", "NON-PARTICIPANT"],
      allowXName: true,
      allowIanaToken: true
    },
    "rsvp": {
      values: ["TRUE", "FALSE"]
    },
    "sent-by": {
      valueType: "cal-address"
    },
    "tzid": {
      matches: /^\//
    },
    "value": {
      // since the value here is a 'type' lowercase is used.
      values: ["binary", "boolean", "cal-address", "date", "date-time",
               "duration", "float", "integer", "period", "recur", "text",
               "time", "uri", "utc-offset"],
      allowXName: true,
      allowIanaToken: true
    }
  };

  // When adding a value here, be sure to add it to the parameter types!
  var icalValues = ICAL.helpers.extend(commonValues, {
    text: createTextType(FROM_ICAL_NEWLINE, TO_ICAL_NEWLINE),

    uri: {
      // TODO
      /* ... */
    },

    "binary": {
      decorate: function(aString) {
        return ICAL.Binary.fromString(aString);
      },

      undecorate: function(aBinary) {
        return aBinary.toString();
      }
    },
    "cal-address": {
      // needs to be an uri
    },
    "date": {
      decorate: function(aValue, aProp) {
        if (design.strict) {
          return ICAL.Time.fromDateString(aValue, aProp);
        } else {
          return ICAL.Time.fromString(aValue, aProp);
        }
      },

      /**
       * undecorates a time object.
       */
      undecorate: function(aValue) {
        return aValue.toString();
      },

      fromICAL: function(aValue) {
        // from: 20120901
        // to: 2012-09-01
        if (!design.strict && aValue.length >= 15) {
          // This is probably a date-time, e.g. 20120901T130000Z
          return icalValues["date-time"].fromICAL(aValue);
        } else {
          return aValue.substr(0, 4) + '-' +
                 aValue.substr(4, 2) + '-' +
                 aValue.substr(6, 2);
        }
      },

      toICAL: function(aValue) {
        // from: 2012-09-01
        // to: 20120901
        var len = aValue.length;

        if (len == 10) {
          return aValue.substr(0, 4) +
                 aValue.substr(5, 2) +
                 aValue.substr(8, 2);
        } else if (len >= 19) {
          return icalValues["date-time"].toICAL(aValue);
        } else {
          //TODO: serialize warning?
          return aValue;
        }

      }
    },
    "date-time": {
      fromICAL: function(aValue) {
        // from: 20120901T130000
        // to: 2012-09-01T13:00:00
        if (!design.strict && aValue.length == 8) {
          // This is probably a date, e.g. 20120901
          return icalValues.date.fromICAL(aValue);
        } else {
          var result = aValue.substr(0, 4) + '-' +
                       aValue.substr(4, 2) + '-' +
                       aValue.substr(6, 2) + 'T' +
                       aValue.substr(9, 2) + ':' +
                       aValue.substr(11, 2) + ':' +
                       aValue.substr(13, 2);

          if (aValue[15] && aValue[15] === 'Z') {
            result += 'Z';
          }

          return result;
        }
      },

      toICAL: function(aValue) {
        // from: 2012-09-01T13:00:00
        // to: 20120901T130000
        var len = aValue.length;

        if (len == 10 && !design.strict) {
          return icalValues.date.toICAL(aValue);
        } else if (len >= 19) {
          var result = aValue.substr(0, 4) +
                       aValue.substr(5, 2) +
                       // grab the (DDTHH) segment
                       aValue.substr(8, 5) +
                       // MM
                       aValue.substr(14, 2) +
                       // SS
                       aValue.substr(17, 2);

          if (aValue[19] && aValue[19] === 'Z') {
            result += 'Z';
          }
          return result;
        } else {
          // TODO: error
          return aValue;
        }
      },

      decorate: function(aValue, aProp) {
        if (design.strict) {
          return ICAL.Time.fromDateTimeString(aValue, aProp);
        } else {
          return ICAL.Time.fromString(aValue, aProp);
        }
      },

      undecorate: function(aValue) {
        return aValue.toString();
      }
    },
    duration: {
      decorate: function(aValue) {
        return ICAL.Duration.fromString(aValue);
      },
      undecorate: function(aValue) {
        return aValue.toString();
      }
    },
    period: {

      fromICAL: function(string) {
        var parts = string.split('/');
        parts[0] = icalValues['date-time'].fromICAL(parts[0]);

        if (!ICAL.Duration.isValueString(parts[1])) {
          parts[1] = icalValues['date-time'].fromICAL(parts[1]);
        }

        return parts;
      },

      toICAL: function(parts) {
        if (!design.strict && parts[0].length == 10) {
          parts[0] = icalValues.date.toICAL(parts[0]);
        } else {
          parts[0] = icalValues['date-time'].toICAL(parts[0]);
        }

        if (!ICAL.Duration.isValueString(parts[1])) {
          if (!design.strict && parts[1].length == 10) {
            parts[1] = icalValues.date.toICAL(parts[1]);
          } else {
            parts[1] = icalValues['date-time'].toICAL(parts[1]);
          }
        }

        return parts.join("/");
      },

      decorate: function(aValue, aProp) {
        return ICAL.Period.fromJSON(aValue, aProp, !design.strict);
      },

      undecorate: function(aValue) {
        return aValue.toJSON();
      }
    },
    recur: {
      fromICAL: function(string) {
        return ICAL.Recur._stringToData(string, true);
      },

      toICAL: function(data) {
        var str = "";
        for (var k in data) {
          /* istanbul ignore if */
          if (!Object.prototype.hasOwnProperty.call(data, k)) {
            continue;
          }
          var val = data[k];
          if (k == "until") {
            if (val.length > 10) {
              val = icalValues['date-time'].toICAL(val);
            } else {
              val = icalValues.date.toICAL(val);
            }
          } else if (k == "wkst") {
            if (typeof val === 'number') {
              val = ICAL.Recur.numericDayToIcalDay(val);
            }
          } else if (Array.isArray(val)) {
            val = val.join(",");
          }
          str += k.toUpperCase() + "=" + val + ";";
        }
        return str.substr(0, str.length - 1);
      },

      decorate: function decorate(aValue) {
        return ICAL.Recur.fromData(aValue);
      },

      undecorate: function(aRecur) {
        return aRecur.toJSON();
      }
    },

    time: {
      fromICAL: function(aValue) {
        // from: MMHHSS(Z)?
        // to: HH:MM:SS(Z)?
        if (aValue.length < 6) {
          // TODO: parser exception?
          return aValue;
        }

        // HH::MM::SSZ?
        var result = aValue.substr(0, 2) + ':' +
                     aValue.substr(2, 2) + ':' +
                     aValue.substr(4, 2);

        if (aValue[6] === 'Z') {
          result += 'Z';
        }

        return result;
      },

      toICAL: function(aValue) {
        // from: HH:MM:SS(Z)?
        // to: MMHHSS(Z)?
        if (aValue.length < 8) {
          //TODO: error
          return aValue;
        }

        var result = aValue.substr(0, 2) +
                     aValue.substr(3, 2) +
                     aValue.substr(6, 2);

        if (aValue[8] === 'Z') {
          result += 'Z';
        }

        return result;
      }
    }
  });

  var icalProperties = ICAL.helpers.extend(commonProperties, {

    "action": DEFAULT_TYPE_TEXT,
    "attach": { defaultType: "uri" },
    "attendee": { defaultType: "cal-address" },
    "calscale": DEFAULT_TYPE_TEXT,
    "class": DEFAULT_TYPE_TEXT,
    "comment": DEFAULT_TYPE_TEXT,
    "completed": DEFAULT_TYPE_DATETIME,
    "contact": DEFAULT_TYPE_TEXT,
    "created": DEFAULT_TYPE_DATETIME,
    "description": DEFAULT_TYPE_TEXT,
    "dtend": DEFAULT_TYPE_DATETIME_DATE,
    "dtstamp": DEFAULT_TYPE_DATETIME,
    "dtstart": DEFAULT_TYPE_DATETIME_DATE,
    "due": DEFAULT_TYPE_DATETIME_DATE,
    "duration": { defaultType: "duration" },
    "exdate": {
      defaultType: "date-time",
      allowedTypes: ["date-time", "date"],
      multiValue: ','
    },
    "exrule": DEFAULT_TYPE_RECUR,
    "freebusy": { defaultType: "period", multiValue: "," },
    "geo": { defaultType: "float", structuredValue: ";" },
    "last-modified": DEFAULT_TYPE_DATETIME,
    "location": DEFAULT_TYPE_TEXT,
    "method": DEFAULT_TYPE_TEXT,
    "organizer": { defaultType: "cal-address" },
    "percent-complete": DEFAULT_TYPE_INTEGER,
    "priority": DEFAULT_TYPE_INTEGER,
    "prodid": DEFAULT_TYPE_TEXT,
    "related-to": DEFAULT_TYPE_TEXT,
    "repeat": DEFAULT_TYPE_INTEGER,
    "rdate": {
      defaultType: "date-time",
      allowedTypes: ["date-time", "date", "period"],
      multiValue: ',',
      detectType: function(string) {
        if (string.indexOf('/') !== -1) {
          return 'period';
        }
        return (string.indexOf('T') === -1) ? 'date' : 'date-time';
      }
    },
    "recurrence-id": DEFAULT_TYPE_DATETIME_DATE,
    "resources": DEFAULT_TYPE_TEXT_MULTI,
    "request-status": DEFAULT_TYPE_TEXT_STRUCTURED,
    "rrule": DEFAULT_TYPE_RECUR,
    "sequence": DEFAULT_TYPE_INTEGER,
    "status": DEFAULT_TYPE_TEXT,
    "summary": DEFAULT_TYPE_TEXT,
    "transp": DEFAULT_TYPE_TEXT,
    "trigger": { defaultType: "duration", allowedTypes: ["duration", "date-time"] },
    "tzoffsetfrom": DEFAULT_TYPE_UTCOFFSET,
    "tzoffsetto": DEFAULT_TYPE_UTCOFFSET,
    "tzurl": DEFAULT_TYPE_URI,
    "tzid": DEFAULT_TYPE_TEXT,
    "tzname": DEFAULT_TYPE_TEXT
  });

  // When adding a value here, be sure to add it to the parameter types!
  var vcardValues = ICAL.helpers.extend(commonValues, {
    text: createTextType(FROM_VCARD_NEWLINE, TO_VCARD_NEWLINE),
    uri: createTextType(FROM_VCARD_NEWLINE, TO_VCARD_NEWLINE),

    date: {
      decorate: function(aValue) {
        return ICAL.VCardTime.fromDateAndOrTimeString(aValue, "date");
      },
      undecorate: function(aValue) {
        return aValue.toString();
      },
      fromICAL: function(aValue) {
        if (aValue.length == 8) {
          return icalValues.date.fromICAL(aValue);
        } else if (aValue[0] == '-' && aValue.length == 6) {
          return aValue.substr(0, 4) + '-' + aValue.substr(4);
        } else {
          return aValue;
        }
      },
      toICAL: function(aValue) {
        if (aValue.length == 10) {
          return icalValues.date.toICAL(aValue);
        } else if (aValue[0] == '-' && aValue.length == 7) {
          return aValue.substr(0, 4) + aValue.substr(5);
        } else {
          return aValue;
        }
      }
    },

    time: {
      decorate: function(aValue) {
        return ICAL.VCardTime.fromDateAndOrTimeString("T" + aValue, "time");
      },
      undecorate: function(aValue) {
        return aValue.toString();
      },
      fromICAL: function(aValue) {
        var splitzone = vcardValues.time._splitZone(aValue, true);
        var zone = splitzone[0], value = splitzone[1];

        //console.log("SPLIT: ",splitzone);

        if (value.length == 6) {
          value = value.substr(0, 2) + ':' +
                  value.substr(2, 2) + ':' +
                  value.substr(4, 2);
        } else if (value.length == 4 && value[0] != '-') {
          value = value.substr(0, 2) + ':' + value.substr(2, 2);
        } else if (value.length == 5) {
          value = value.substr(0, 3) + ':' + value.substr(3, 2);
        }

        if (zone.length == 5 && (zone[0] == '-' || zone[0] == '+')) {
          zone = zone.substr(0, 3) + ':' + zone.substr(3);
        }

        return value + zone;
      },

      toICAL: function(aValue) {
        var splitzone = vcardValues.time._splitZone(aValue);
        var zone = splitzone[0], value = splitzone[1];

        if (value.length == 8) {
          value = value.substr(0, 2) +
                  value.substr(3, 2) +
                  value.substr(6, 2);
        } else if (value.length == 5 && value[0] != '-') {
          value = value.substr(0, 2) + value.substr(3, 2);
        } else if (value.length == 6) {
          value = value.substr(0, 3) + value.substr(4, 2);
        }

        if (zone.length == 6 && (zone[0] == '-' || zone[0] == '+')) {
          zone = zone.substr(0, 3) + zone.substr(4);
        }

        return value + zone;
      },

      _splitZone: function(aValue, isFromIcal) {
        var lastChar = aValue.length - 1;
        var signChar = aValue.length - (isFromIcal ? 5 : 6);
        var sign = aValue[signChar];
        var zone, value;

        if (aValue[lastChar] == 'Z') {
          zone = aValue[lastChar];
          value = aValue.substr(0, lastChar);
        } else if (aValue.length > 6 && (sign == '-' || sign == '+')) {
          zone = aValue.substr(signChar);
          value = aValue.substr(0, signChar);
        } else {
          zone = "";
          value = aValue;
        }

        return [zone, value];
      }
    },

    "date-time": {
      decorate: function(aValue) {
        return ICAL.VCardTime.fromDateAndOrTimeString(aValue, "date-time");
      },

      undecorate: function(aValue) {
        return aValue.toString();
      },

      fromICAL: function(aValue) {
        return vcardValues['date-and-or-time'].fromICAL(aValue);
      },

      toICAL: function(aValue) {
        return vcardValues['date-and-or-time'].toICAL(aValue);
      }
    },

    "date-and-or-time": {
      decorate: function(aValue) {
        return ICAL.VCardTime.fromDateAndOrTimeString(aValue, "date-and-or-time");
      },

      undecorate: function(aValue) {
        return aValue.toString();
      },

      fromICAL: function(aValue) {
        var parts = aValue.split('T');
        return (parts[0] ? vcardValues.date.fromICAL(parts[0]) : '') +
               (parts[1] ? 'T' + vcardValues.time.fromICAL(parts[1]) : '');
      },

      toICAL: function(aValue) {
        var parts = aValue.split('T');
        return vcardValues.date.toICAL(parts[0]) +
               (parts[1] ? 'T' + vcardValues.time.toICAL(parts[1]) : '');

      }
    },
    timestamp: icalValues['date-time'],
    "language-tag": {
      matches: /^[a-zA-Z0-9-]+$/ // Could go with a more strict regex here
    }
  });

  var vcardParams = {
    "type": {
      valueType: "text",
      multiValue: ","
    },
    "value": {
      // since the value here is a 'type' lowercase is used.
      values: ["text", "uri", "date", "time", "date-time", "date-and-or-time",
               "timestamp", "boolean", "integer", "float", "utc-offset",
               "language-tag"],
      allowXName: true,
      allowIanaToken: true
    }
  };

  var vcardProperties = ICAL.helpers.extend(commonProperties, {
    "adr": { defaultType: "text", structuredValue: ";", multiValue: "," },
    "anniversary": DEFAULT_TYPE_DATE_ANDOR_TIME,
    "bday": DEFAULT_TYPE_DATE_ANDOR_TIME,
    "caladruri": DEFAULT_TYPE_URI,
    "caluri": DEFAULT_TYPE_URI,
    "clientpidmap": DEFAULT_TYPE_TEXT_STRUCTURED,
    "email": DEFAULT_TYPE_TEXT,
    "fburl": DEFAULT_TYPE_URI,
    "fn": DEFAULT_TYPE_TEXT,
    "gender": DEFAULT_TYPE_TEXT_STRUCTURED,
    "geo": DEFAULT_TYPE_URI,
    "impp": DEFAULT_TYPE_URI,
    "key": DEFAULT_TYPE_URI,
    "kind": DEFAULT_TYPE_TEXT,
    "lang": { defaultType: "language-tag" },
    "logo": DEFAULT_TYPE_URI,
    "member": DEFAULT_TYPE_URI,
    "n": { defaultType: "text", structuredValue: ";", multiValue: "," },
    "nickname": DEFAULT_TYPE_TEXT_MULTI,
    "note": DEFAULT_TYPE_TEXT,
    "org": { defaultType: "text", structuredValue: ";" },
    "photo": DEFAULT_TYPE_URI,
    "related": DEFAULT_TYPE_URI,
    "rev": { defaultType: "timestamp" },
    "role": DEFAULT_TYPE_TEXT,
    "sound": DEFAULT_TYPE_URI,
    "source": DEFAULT_TYPE_URI,
    "tel": { defaultType: "uri", allowedTypes: ["uri", "text"] },
    "title": DEFAULT_TYPE_TEXT,
    "tz": { defaultType: "text", allowedTypes: ["text", "utc-offset", "uri"] },
    "xml": DEFAULT_TYPE_TEXT
  });

  var vcard3Values = ICAL.helpers.extend(commonValues, {
    binary: icalValues.binary,
    date: vcardValues.date,
    "date-time": vcardValues["date-time"],
    "phone-number": {
      // TODO
      /* ... */
    },
    uri: icalValues.uri,
    text: icalValues.text,
    time: icalValues.time,
    vcard: icalValues.text,
    "utc-offset": {
      toICAL: function(aValue) {
        return aValue.substr(0, 7);
      },

      fromICAL: function(aValue) {
        return aValue.substr(0, 7);
      },

      decorate: function(aValue) {
        return ICAL.UtcOffset.fromString(aValue);
      },

      undecorate: function(aValue) {
        return aValue.toString();
      }
    }
  });

  var vcard3Params = {
    "type": {
      valueType: "text",
      multiValue: ","
    },
    "value": {
      // since the value here is a 'type' lowercase is used.
      values: ["text", "uri", "date", "date-time", "phone-number", "time",
               "boolean", "integer", "float", "utc-offset", "vcard", "binary"],
      allowXName: true,
      allowIanaToken: true
    }
  };

  var vcard3Properties = ICAL.helpers.extend(commonProperties, {
    fn: DEFAULT_TYPE_TEXT,
    n: { defaultType: "text", structuredValue: ";", multiValue: "," },
    nickname: DEFAULT_TYPE_TEXT_MULTI,
    photo: { defaultType: "binary", allowedTypes: ["binary", "uri"] },
    bday: {
      defaultType: "date-time",
      allowedTypes: ["date-time", "date"],
      detectType: function(string) {
        return (string.indexOf('T') === -1) ? 'date' : 'date-time';
      }
    },

    adr: { defaultType: "text", structuredValue: ";", multiValue: "," },
    label: DEFAULT_TYPE_TEXT,

    tel: { defaultType: "phone-number" },
    email: DEFAULT_TYPE_TEXT,
    mailer: DEFAULT_TYPE_TEXT,

    tz: { defaultType: "utc-offset", allowedTypes: ["utc-offset", "text"] },
    geo: { defaultType: "float", structuredValue: ";" },

    title: DEFAULT_TYPE_TEXT,
    role: DEFAULT_TYPE_TEXT,
    logo: { defaultType: "binary", allowedTypes: ["binary", "uri"] },
    agent: { defaultType: "vcard", allowedTypes: ["vcard", "text", "uri"] },
    org: DEFAULT_TYPE_TEXT_STRUCTURED,

    note: DEFAULT_TYPE_TEXT_MULTI,
    prodid: DEFAULT_TYPE_TEXT,
    rev: {
      defaultType: "date-time",
      allowedTypes: ["date-time", "date"],
      detectType: function(string) {
        return (string.indexOf('T') === -1) ? 'date' : 'date-time';
      }
    },
    "sort-string": DEFAULT_TYPE_TEXT,
    sound: { defaultType: "binary", allowedTypes: ["binary", "uri"] },

    class: DEFAULT_TYPE_TEXT,
    key: { defaultType: "binary", allowedTypes: ["binary", "text"] }
  });

  /**
   * iCalendar design set
   * @type {ICAL.design.designSet}
   */
  var icalSet = {
    value: icalValues,
    param: icalParams,
    property: icalProperties
  };

  /**
   * vCard 4.0 design set
   * @type {ICAL.design.designSet}
   */
  var vcardSet = {
    value: vcardValues,
    param: vcardParams,
    property: vcardProperties
  };

  /**
   * vCard 3.0 design set
   * @type {ICAL.design.designSet}
   */
  var vcard3Set = {
    value: vcard3Values,
    param: vcard3Params,
    property: vcard3Properties
  };

  /**
   * The design data, used by the parser to determine types for properties and
   * other metadata needed to produce correct jCard/jCal data.
   *
   * @alias ICAL.design
   * @namespace
   */
  var design = {
    /**
     * A designSet describes value, parameter and property data. It is used by
     * ther parser and stringifier in components and properties to determine they
     * should be represented.
     *
     * @typedef {Object} designSet
     * @memberOf ICAL.design
     * @property {Object} value       Definitions for value types, keys are type names
     * @property {Object} param       Definitions for params, keys are param names
     * @property {Object} property    Defintions for properties, keys are property names
     */

    /**
     * Can be set to false to make the parser more lenient.
     */
    strict: false,

    /**
     * The default set for new properties and components if none is specified.
     * @type {ICAL.design.designSet}
     */
    defaultSet: icalSet,

    /**
     * The default type for unknown properties
     * @type {String}
     */
    defaultType: 'unknown',

    /**
     * Holds the design set for known top-level components
     *
     * @type {Object}
     * @property {ICAL.design.designSet} vcard       vCard VCARD
     * @property {ICAL.design.designSet} vevent      iCalendar VEVENT
     * @property {ICAL.design.designSet} vtodo       iCalendar VTODO
     * @property {ICAL.design.designSet} vjournal    iCalendar VJOURNAL
     * @property {ICAL.design.designSet} valarm      iCalendar VALARM
     * @property {ICAL.design.designSet} vtimezone   iCalendar VTIMEZONE
     * @property {ICAL.design.designSet} daylight    iCalendar DAYLIGHT
     * @property {ICAL.design.designSet} standard    iCalendar STANDARD
     *
     * @example
     * var propertyName = 'fn';
     * var componentDesign = ICAL.design.components.vcard;
     * var propertyDetails = componentDesign.property[propertyName];
     * if (propertyDetails.defaultType == 'text') {
     *   // Yep, sure is...
     * }
     */
    components: {
      vcard: vcardSet,
      vcard3: vcard3Set,
      vevent: icalSet,
      vtodo: icalSet,
      vjournal: icalSet,
      valarm: icalSet,
      vtimezone: icalSet,
      daylight: icalSet,
      standard: icalSet
    },


    /**
     * The design set for iCalendar (rfc5545/rfc7265) components.
     * @type {ICAL.design.designSet}
     */
    icalendar: icalSet,

    /**
     * The design set for vCard (rfc6350/rfc7095) components.
     * @type {ICAL.design.designSet}
     */
    vcard: vcardSet,

    /**
     * The design set for vCard (rfc2425/rfc2426/rfc7095) components.
     * @type {ICAL.design.designSet}
     */
    vcard3: vcard3Set,

    /**
     * Gets the design set for the given component name.
     *
     * @param {String} componentName        The name of the component
     * @return {ICAL.design.designSet}      The design set for the component
     */
    getDesignSet: function(componentName) {
      var isInDesign = componentName && componentName in design.components;
      return isInDesign ? design.components[componentName] : design.defaultSet;
    }
  };

  return design;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * Contains various functions to convert jCal and jCard data back into
 * iCalendar and vCard.
 * @namespace
 */
ICAL.stringify = (function() {
  'use strict';

  var LINE_ENDING = '\r\n';
  var DEFAULT_VALUE_TYPE = 'unknown';

  var design = ICAL.design;
  var helpers = ICAL.helpers;

  /**
   * Convert a full jCal/jCard array into a iCalendar/vCard string.
   *
   * @function ICAL.stringify
   * @variation function
   * @param {Array} jCal    The jCal/jCard document
   * @return {String}       The stringified iCalendar/vCard document
   */
  function stringify(jCal) {
    if (typeof jCal[0] == "string") {
      // This is a single component
      jCal = [jCal];
    }

    var i = 0;
    var len = jCal.length;
    var result = '';

    for (; i < len; i++) {
      result += stringify.component(jCal[i]) + LINE_ENDING;
    }

    return result;
  }

  /**
   * Converts an jCal component array into a ICAL string.
   * Recursive will resolve sub-components.
   *
   * Exact component/property order is not saved all
   * properties will come before subcomponents.
   *
   * @function ICAL.stringify.component
   * @param {Array} component
   *        jCal/jCard fragment of a component
   * @param {ICAL.design.designSet} designSet
   *        The design data to use for this component
   * @return {String}       The iCalendar/vCard string
   */
  stringify.component = function(component, designSet) {
    var name = component[0].toUpperCase();
    var result = 'BEGIN:' + name + LINE_ENDING;

    var props = component[1];
    var propIdx = 0;
    var propLen = props.length;

    var designSetName = component[0];
    // rfc6350 requires that in vCard 4.0 the first component is the VERSION
    // component with as value 4.0, note that 3.0 does not have this requirement.
    if (designSetName === 'vcard' && component[1].length > 0 &&
            !(component[1][0][0] === "version" && component[1][0][3] === "4.0")) {
      designSetName = "vcard3";
    }
    designSet = designSet || design.getDesignSet(designSetName);

    for (; propIdx < propLen; propIdx++) {
      result += stringify.property(props[propIdx], designSet) + LINE_ENDING;
    }

    // Ignore subcomponents if none exist, e.g. in vCard.
    var comps = component[2] || [];
    var compIdx = 0;
    var compLen = comps.length;

    for (; compIdx < compLen; compIdx++) {
      result += stringify.component(comps[compIdx], designSet) + LINE_ENDING;
    }

    result += 'END:' + name;
    return result;
  };

  /**
   * Converts a single jCal/jCard property to a iCalendar/vCard string.
   *
   * @function ICAL.stringify.property
   * @param {Array} property
   *        jCal/jCard property array
   * @param {ICAL.design.designSet} designSet
   *        The design data to use for this property
   * @param {Boolean} noFold
   *        If true, the line is not folded
   * @return {String}       The iCalendar/vCard string
   */
  stringify.property = function(property, designSet, noFold) {
    var name = property[0].toUpperCase();
    var jsName = property[0];
    var params = property[1];

    var line = name;

    var paramName;
    for (paramName in params) {
      var value = params[paramName];

      /* istanbul ignore else */
      if (params.hasOwnProperty(paramName)) {
        var multiValue = (paramName in designSet.param) && designSet.param[paramName].multiValue;
        if (multiValue && Array.isArray(value)) {
          if (designSet.param[paramName].multiValueSeparateDQuote) {
            multiValue = '"' + multiValue + '"';
          }
          value = value.map(stringify._rfc6868Unescape);
          value = stringify.multiValue(value, multiValue, "unknown", null, designSet);
        } else {
          value = stringify._rfc6868Unescape(value);
        }


        line += ';' + paramName.toUpperCase();
        line += '=' + stringify.propertyValue(value);
      }
    }

    if (property.length === 3) {
      // If there are no values, we must assume a blank value
      return line + ':';
    }

    var valueType = property[2];

    if (!designSet) {
      designSet = design.defaultSet;
    }

    var propDetails;
    var multiValue = false;
    var structuredValue = false;
    var isDefault = false;

    if (jsName in designSet.property) {
      propDetails = designSet.property[jsName];

      if ('multiValue' in propDetails) {
        multiValue = propDetails.multiValue;
      }

      if (('structuredValue' in propDetails) && Array.isArray(property[3])) {
        structuredValue = propDetails.structuredValue;
      }

      if ('defaultType' in propDetails) {
        if (valueType === propDetails.defaultType) {
          isDefault = true;
        }
      } else {
        if (valueType === DEFAULT_VALUE_TYPE) {
          isDefault = true;
        }
      }
    } else {
      if (valueType === DEFAULT_VALUE_TYPE) {
        isDefault = true;
      }
    }

    // push the VALUE property if type is not the default
    // for the current property.
    if (!isDefault) {
      // value will never contain ;/:/, so we don't escape it here.
      line += ';VALUE=' + valueType.toUpperCase();
    }

    line += ':';

    if (multiValue && structuredValue) {
      line += stringify.multiValue(
        property[3], structuredValue, valueType, multiValue, designSet, structuredValue
      );
    } else if (multiValue) {
      line += stringify.multiValue(
        property.slice(3), multiValue, valueType, null, designSet, false
      );
    } else if (structuredValue) {
      line += stringify.multiValue(
        property[3], structuredValue, valueType, null, designSet, structuredValue
      );
    } else {
      line += stringify.value(property[3], valueType, designSet, false);
    }

    return noFold ? line : ICAL.helpers.foldline(line);
  };

  /**
   * Handles escaping of property values that may contain:
   *
   *    COLON (:), SEMICOLON (;), or COMMA (,)
   *
   * If any of the above are present the result is wrapped
   * in double quotes.
   *
   * @function ICAL.stringify.propertyValue
   * @param {String} value      Raw property value
   * @return {String}           Given or escaped value when needed
   */
  stringify.propertyValue = function(value) {

    if ((helpers.unescapedIndexOf(value, ',') === -1) &&
        (helpers.unescapedIndexOf(value, ':') === -1) &&
        (helpers.unescapedIndexOf(value, ';') === -1)) {

      return value;
    }

    return '"' + value + '"';
  };

  /**
   * Converts an array of ical values into a single
   * string based on a type and a delimiter value (like ",").
   *
   * @function ICAL.stringify.multiValue
   * @param {Array} values      List of values to convert
   * @param {String} delim      Used to join the values (",", ";", ":")
   * @param {String} type       Lowecase ical value type
   *        (like boolean, date-time, etc..)
   * @param {?String} innerMulti If set, each value will again be processed
   *        Used for structured values
   * @param {ICAL.design.designSet} designSet
   *        The design data to use for this property
   *
   * @return {String}           iCalendar/vCard string for value
   */
  stringify.multiValue = function(values, delim, type, innerMulti, designSet, structuredValue) {
    var result = '';
    var len = values.length;
    var i = 0;

    for (; i < len; i++) {
      if (innerMulti && Array.isArray(values[i])) {
        result += stringify.multiValue(values[i], innerMulti, type, null, designSet, structuredValue);
      } else {
        result += stringify.value(values[i], type, designSet, structuredValue);
      }

      if (i !== (len - 1)) {
        result += delim;
      }
    }

    return result;
  };

  /**
   * Processes a single ical value runs the associated "toICAL" method from the
   * design value type if available to convert the value.
   *
   * @function ICAL.stringify.value
   * @param {String|Number} value       A formatted value
   * @param {String} type               Lowercase iCalendar/vCard value type
   *  (like boolean, date-time, etc..)
   * @return {String}                   iCalendar/vCard value for single value
   */
  stringify.value = function(value, type, designSet, structuredValue) {
    if (type in designSet.value && 'toICAL' in designSet.value[type]) {
      return designSet.value[type].toICAL(value, structuredValue);
    }
    return value;
  };

  /**
   * Internal helper for rfc6868. Exposing this on ICAL.stringify so that
   * hackers can disable the rfc6868 parsing if the really need to.
   *
   * @param {String} val        The value to unescape
   * @return {String}           The escaped value
   */
  stringify._rfc6868Unescape = function(val) {
    return val.replace(/[\n^"]/g, function(x) {
      return RFC6868_REPLACE_MAP[x];
    });
  };
  var RFC6868_REPLACE_MAP = { '"': "^'", "\n": "^n", "^": "^^" };

  return stringify;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * Contains various functions to parse iCalendar and vCard data.
 * @namespace
 */
ICAL.parse = (function() {
  'use strict';

  var CHAR = /[^ \t]/;
  var MULTIVALUE_DELIMITER = ',';
  var VALUE_DELIMITER = ':';
  var PARAM_DELIMITER = ';';
  var PARAM_NAME_DELIMITER = '=';
  var DEFAULT_VALUE_TYPE = 'unknown';
  var DEFAULT_PARAM_TYPE = 'text';

  var design = ICAL.design;
  var helpers = ICAL.helpers;

  /**
   * An error that occurred during parsing.
   *
   * @param {String} message        The error message
   * @memberof ICAL.parse
   * @extends {Error}
   * @class
   */
  function ParserError(message) {
    this.message = message;
    this.name = 'ParserError';

    try {
      throw new Error();
    } catch (e) {
      if (e.stack) {
        var split = e.stack.split('\n');
        split.shift();
        this.stack = split.join('\n');
      }
    }
  }

  ParserError.prototype = Error.prototype;

  /**
   * Parses iCalendar or vCard data into a raw jCal object. Consult
   * documentation on the {@tutorial layers|layers of parsing} for more
   * details.
   *
   * @function ICAL.parse
   * @variation function
   * @todo Fix the API to be more clear on the return type
   * @param {String} input      The string data to parse
   * @return {Object|Object[]}  A single jCal object, or an array thereof
   */
  function parser(input) {
    var state = {};
    var root = state.component = [];

    state.stack = [root];

    parser._eachLine(input, function(err, line) {
      parser._handleContentLine(line, state);
    });


    // when there are still items on the stack
    // throw a fatal error, a component was not closed
    // correctly in that case.
    if (state.stack.length > 1) {
      throw new ParserError(
        'invalid ical body. component began but did not end'
      );
    }

    state = null;

    return (root.length == 1 ? root[0] : root);
  }

  /**
   * Parse an iCalendar property value into the jCal for a single property
   *
   * @function ICAL.parse.property
   * @param {String} str
   *   The iCalendar property string to parse
   * @param {ICAL.design.designSet=} designSet
   *   The design data to use for this property
   * @return {Object}
   *   The jCal Object containing the property
   */
  parser.property = function(str, designSet) {
    var state = {
      component: [[], []],
      designSet: designSet || design.defaultSet
    };
    parser._handleContentLine(str, state);
    return state.component[1][0];
  };

  /**
   * Convenience method to parse a component. You can use ICAL.parse() directly
   * instead.
   *
   * @function ICAL.parse.component
   * @see ICAL.parse(function)
   * @param {String} str    The iCalendar component string to parse
   * @return {Object}       The jCal Object containing the component
   */
  parser.component = function(str) {
    return parser(str);
  };

  // classes & constants
  parser.ParserError = ParserError;

  /**
   * The state for parsing content lines from an iCalendar/vCard string.
   *
   * @private
   * @memberof ICAL.parse
   * @typedef {Object} parserState
   * @property {ICAL.design.designSet} designSet    The design set to use for parsing
   * @property {ICAL.Component[]} stack             The stack of components being processed
   * @property {ICAL.Component} component           The currently active component
   */


  /**
   * Handles a single line of iCalendar/vCard, updating the state.
   *
   * @private
   * @function ICAL.parse._handleContentLine
   * @param {String} line               The content line to process
   * @param {ICAL.parse.parserState}    The current state of the line parsing
   */
  parser._handleContentLine = function(line, state) {
    // break up the parts of the line
    var valuePos = line.indexOf(VALUE_DELIMITER);
    var paramPos = line.indexOf(PARAM_DELIMITER);

    var lastParamIndex;
    var lastValuePos;

    // name of property or begin/end
    var name;
    var value;
    // params is only overridden if paramPos !== -1.
    // we can't do params = params || {} later on
    // because it sacrifices ops.
    var params = {};

    /**
     * Different property cases
     *
     *
     * 1. RRULE:FREQ=foo
     *    // FREQ= is not a param but the value
     *
     * 2. ATTENDEE;ROLE=REQ-PARTICIPANT;
     *    // ROLE= is a param because : has not happened yet
     */
      // when the parameter delimiter is after the
      // value delimiter then its not a parameter.

    if ((paramPos !== -1 && valuePos !== -1)) {
      // when the parameter delimiter is after the
      // value delimiter then its not a parameter.
      if (paramPos > valuePos) {
        paramPos = -1;
      }
    }

    var parsedParams;
    if (paramPos !== -1) {
      name = line.substring(0, paramPos).toLowerCase();
      parsedParams = parser._parseParameters(line.substring(paramPos), 0, state.designSet);
      if (parsedParams[2] == -1) {
        throw new ParserError("Invalid parameters in '" + line + "'");
      }
      params = parsedParams[0];
      lastParamIndex = parsedParams[1].length + parsedParams[2] + paramPos;
      if ((lastValuePos =
        line.substring(lastParamIndex).indexOf(VALUE_DELIMITER)) !== -1) {
        value = line.substring(lastParamIndex + lastValuePos + 1);
      } else {
        throw new ParserError("Missing parameter value in '" + line + "'");
      }
    } else if (valuePos !== -1) {
      // without parmeters (BEGIN:VCAENDAR, CLASS:PUBLIC)
      name = line.substring(0, valuePos).toLowerCase();
      value = line.substring(valuePos + 1);

      if (name === 'begin') {
        var newComponent = [value.toLowerCase(), [], []];
        if (state.stack.length === 1) {
          state.component.push(newComponent);
        } else {
          state.component[2].push(newComponent);
        }
        state.stack.push(state.component);
        state.component = newComponent;
        if (!state.designSet) {
          state.designSet = design.getDesignSet(state.component[0]);
        }
        return;
      } else if (name === 'end') {
        state.component = state.stack.pop();
        return;
      }
      // If its not begin/end, then this is a property with an empty value,
      // which should be considered valid.
    } else {
      /**
       * Invalid line.
       * The rational to throw an error is we will
       * never be certain that the rest of the file
       * is sane and its unlikely that we can serialize
       * the result correctly either.
       */
      Logger.log('invalid line (no token ";" or ":") "' + line + '"');
      return;
    }

    var valueType;
    var multiValue = false;
    var structuredValue = false;
    var propertyDetails;

    if (name in state.designSet.property) {
      propertyDetails = state.designSet.property[name];

      if ('multiValue' in propertyDetails) {
        multiValue = propertyDetails.multiValue;
      }

      if ('structuredValue' in propertyDetails) {
        structuredValue = propertyDetails.structuredValue;
      }

      if (value && 'detectType' in propertyDetails) {
        valueType = propertyDetails.detectType(value);
      }
    }

    // attempt to determine value
    if (!valueType) {
      if (!('value' in params)) {
        if (propertyDetails) {
          valueType = propertyDetails.defaultType;
        } else {
          valueType = DEFAULT_VALUE_TYPE;
        }
      } else {
        // possible to avoid this?
        valueType = params.value.toLowerCase();
      }
    }

    delete params.value;

    /**
     * Note on `var result` juggling:
     *
     * I observed that building the array in pieces has adverse
     * effects on performance, so where possible we inline the creation.
     * Its a little ugly but resulted in ~2000 additional ops/sec.
     */

    var result;
    if (multiValue && structuredValue) {
      value = parser._parseMultiValue(value, structuredValue, valueType, [], multiValue, state.designSet, structuredValue);
      result = [name, params, valueType, value];
    } else if (multiValue) {
      result = [name, params, valueType];
      parser._parseMultiValue(value, multiValue, valueType, result, null, state.designSet, false);
    } else if (structuredValue) {
      value = parser._parseMultiValue(value, structuredValue, valueType, [], null, state.designSet, structuredValue);
      result = [name, params, valueType, value];
    } else {
      value = parser._parseValue(value, valueType, state.designSet, false);
      result = [name, params, valueType, value];
    }
    // rfc6350 requires that in vCard 4.0 the first component is the VERSION
    // component with as value 4.0, note that 3.0 does not have this requirement.
    if (state.component[0] === 'vcard' && state.component[1].length === 0 &&
            !(name === 'version' && value === '4.0')) {
      state.designSet = design.getDesignSet("vcard3");
    }
    state.component[1].push(result);
  };

  /**
   * Parse a value from the raw value into the jCard/jCal value.
   *
   * @private
   * @function ICAL.parse._parseValue
   * @param {String} value          Original value
   * @param {String} type           Type of value
   * @param {Object} designSet      The design data to use for this value
   * @return {Object} varies on type
   */
  parser._parseValue = function(value, type, designSet, structuredValue) {
    if (type in designSet.value && 'fromICAL' in designSet.value[type]) {
      return designSet.value[type].fromICAL(value, structuredValue);
    }
    return value;
  };

  /**
   * Parse parameters from a string to object.
   *
   * @function ICAL.parse._parseParameters
   * @private
   * @param {String} line           A single unfolded line
   * @param {Numeric} start         Position to start looking for properties
   * @param {Object} designSet      The design data to use for this property
   * @return {Object} key/value pairs
   */
  parser._parseParameters = function(line, start, designSet) {
    var lastParam = start;
    var pos = 0;
    var delim = PARAM_NAME_DELIMITER;
    var result = {};
    var name, lcname;
    var value, valuePos = -1;
    var type, multiValue, mvdelim;

    // find the next '=' sign
    // use lastParam and pos to find name
    // check if " is used if so get value from "->"
    // then increment pos to find next ;

    while ((pos !== false) &&
           (pos = helpers.unescapedIndexOf(line, delim, pos + 1)) !== -1) {

      name = line.substr(lastParam + 1, pos - lastParam - 1);
      if (name.length == 0) {
        throw new ParserError("Empty parameter name in '" + line + "'");
      }
      lcname = name.toLowerCase();
      mvdelim = false;
      multiValue = false;

      if (lcname in designSet.param && designSet.param[lcname].valueType) {
        type = designSet.param[lcname].valueType;
      } else {
        type = DEFAULT_PARAM_TYPE;
      }

      if (lcname in designSet.param) {
        multiValue = designSet.param[lcname].multiValue;
        if (designSet.param[lcname].multiValueSeparateDQuote) {
          mvdelim = parser._rfc6868Escape('"' + multiValue + '"');
        }
      }

      var nextChar = line[pos + 1];
      if (nextChar === '"') {
        valuePos = pos + 2;
        pos = helpers.unescapedIndexOf(line, '"', valuePos);
        if (multiValue && pos != -1) {
            var extendedValue = true;
            while (extendedValue) {
              if (line[pos + 1] == multiValue && line[pos + 2] == '"') {
                pos = helpers.unescapedIndexOf(line, '"', pos + 3);
              } else {
                extendedValue = false;
              }
            }
          }
        if (pos === -1) {
          throw new ParserError(
            'invalid line (no matching double quote) "' + line + '"'
          );
        }
        value = line.substr(valuePos, pos - valuePos);
        lastParam = helpers.unescapedIndexOf(line, PARAM_DELIMITER, pos);
        if (lastParam === -1) {
          pos = false;
        }
      } else {
        valuePos = pos + 1;

        // move to next ";"
        var nextPos = helpers.unescapedIndexOf(line, PARAM_DELIMITER, valuePos);
        var propValuePos = helpers.unescapedIndexOf(line, VALUE_DELIMITER, valuePos);
        if (propValuePos !== -1 && nextPos > propValuePos) {
          // this is a delimiter in the property value, let's stop here
          nextPos = propValuePos;
          pos = false;
        } else if (nextPos === -1) {
          // no ";"
          if (propValuePos === -1) {
            nextPos = line.length;
          } else {
            nextPos = propValuePos;
          }
          pos = false;
        } else {
          lastParam = nextPos;
          pos = nextPos;
        }

        value = line.substr(valuePos, nextPos - valuePos);
      }

      value = parser._rfc6868Escape(value);
      if (multiValue) {
        var delimiter = mvdelim || multiValue;
        value = parser._parseMultiValue(value, delimiter, type, [], null, designSet);
      } else {
        value = parser._parseValue(value, type, designSet);
      }

      if (multiValue && (lcname in result)) {
        if (Array.isArray(result[lcname])) {
          result[lcname].push(value);
        } else {
          result[lcname] = [
            result[lcname],
            value
          ];
        }
      } else {
        result[lcname] = value;
      }
    }
    return [result, value, valuePos];
  };

  /**
   * Internal helper for rfc6868. Exposing this on ICAL.parse so that
   * hackers can disable the rfc6868 parsing if the really need to.
   *
   * @function ICAL.parse._rfc6868Escape
   * @param {String} val        The value to escape
   * @return {String}           The escaped value
   */
  parser._rfc6868Escape = function(val) {
    return val.replace(/\^['n^]/g, function(x) {
      return RFC6868_REPLACE_MAP[x];
    });
  };
  var RFC6868_REPLACE_MAP = { "^'": '"', "^n": "\n", "^^": "^" };

  /**
   * Parse a multi value string. This function is used either for parsing
   * actual multi-value property's values, or for handling parameter values. It
   * can be used for both multi-value properties and structured value properties.
   *
   * @private
   * @function ICAL.parse._parseMultiValue
   * @param {String} buffer     The buffer containing the full value
   * @param {String} delim      The multi-value delimiter
   * @param {String} type       The value type to be parsed
   * @param {Array.<?>} result        The array to append results to, varies on value type
   * @param {String} innerMulti The inner delimiter to split each value with
   * @param {ICAL.design.designSet} designSet   The design data for this value
   * @return {?|Array.<?>}            Either an array of results, or the first result
   */
  parser._parseMultiValue = function(buffer, delim, type, result, innerMulti, designSet, structuredValue) {
    var pos = 0;
    var lastPos = 0;
    var value;
    if (delim.length === 0) {
      return buffer;
    }

    // split each piece
    while ((pos = helpers.unescapedIndexOf(buffer, delim, lastPos)) !== -1) {
      value = buffer.substr(lastPos, pos - lastPos);
      if (innerMulti) {
        value = parser._parseMultiValue(value, innerMulti, type, [], null, designSet, structuredValue);
      } else {
        value = parser._parseValue(value, type, designSet, structuredValue);
      }
      result.push(value);
      lastPos = pos + delim.length;
    }

    // on the last piece take the rest of string
    value = buffer.substr(lastPos);
    if (innerMulti) {
      value = parser._parseMultiValue(value, innerMulti, type, [], null, designSet, structuredValue);
    } else {
      value = parser._parseValue(value, type, designSet, structuredValue);
    }
    result.push(value);

    return result.length == 1 ? result[0] : result;
  };

  /**
   * Process a complete buffer of iCalendar/vCard data line by line, correctly
   * unfolding content. Each line will be processed with the given callback
   *
   * @private
   * @function ICAL.parse._eachLine
   * @param {String} buffer                         The buffer to process
   * @param {function(?String, String)} callback    The callback for each line
   */
  parser._eachLine = function(buffer, callback) {
    var len = buffer.length;
    var lastPos = buffer.search(CHAR);
    var pos = lastPos;
    var line;
    var firstChar;

    var newlineOffset;

    do {
      pos = buffer.indexOf('\n', lastPos) + 1;

      if (pos > 1 && buffer[pos - 2] === '\r') {
        newlineOffset = 2;
      } else {
        newlineOffset = 1;
      }

      if (pos === 0) {
        pos = len;
        newlineOffset = 0;
      }

      firstChar = buffer[lastPos];

      if (firstChar === ' ' || firstChar === '\t') {
        // add to line
        line += buffer.substr(
          lastPos + 1,
          pos - lastPos - (newlineOffset + 1)
        );
      } else {
        if (line)
          callback(null, line);
        // push line
        line = buffer.substr(
          lastPos,
          pos - lastPos - newlineOffset
        );
      }

      lastPos = pos;
    } while (pos !== len);

    // extra ending line
    line = line.trim();

    if (line.length)
      callback(null, line);
  };

  return parser;

}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.Component = (function() {
  'use strict';

  var PROPERTY_INDEX = 1;
  var COMPONENT_INDEX = 2;
  var NAME_INDEX = 0;

  /**
   * @classdesc
   * Wraps a jCal component, adding convenience methods to add, remove and
   * update subcomponents and properties.
   *
   * @class
   * @alias ICAL.Component
   * @param {Array|String} jCal         Raw jCal component data OR name of new
   *                                      component
   * @param {ICAL.Component} parent     Parent component to associate
   */
  function Component(jCal, parent) {
    if (typeof(jCal) === 'string') {
      // jCal spec (name, properties, components)
      jCal = [jCal, [], []];
    }

    // mostly for legacy reasons.
    this.jCal = jCal;

    this.parent = parent || null;
  }

  Component.prototype = {
    /**
     * Hydrated properties are inserted into the _properties array at the same
     * position as in the jCal array, so its possible the array contains
     * undefined values for unhydrdated properties. To avoid iterating the
     * array when checking if all properties have been hydrated, we save the
     * count here.
     *
     * @type {Number}
     * @private
     */
    _hydratedPropertyCount: 0,

    /**
     * The same count as for _hydratedPropertyCount, but for subcomponents
     *
     * @type {Number}
     * @private
     */
    _hydratedComponentCount: 0,

    /**
     * The name of this component
     * @readonly
     */
    get name() {
      return this.jCal[NAME_INDEX];
    },

    /**
     * The design set for this component, e.g. icalendar vs vcard
     *
     * @type {ICAL.design.designSet}
     * @private
     */
    get _designSet() {
      var parentDesign = this.parent && this.parent._designSet;
      return parentDesign || ICAL.design.getDesignSet(this.name);
    },

    _hydrateComponent: function(index) {
      if (!this._components) {
        this._components = [];
        this._hydratedComponentCount = 0;
      }

      if (this._components[index]) {
        return this._components[index];
      }

      var comp = new Component(
        this.jCal[COMPONENT_INDEX][index],
        this
      );

      this._hydratedComponentCount++;
      return (this._components[index] = comp);
    },

    _hydrateProperty: function(index) {
      if (!this._properties) {
        this._properties = [];
        this._hydratedPropertyCount = 0;
      }

      if (this._properties[index]) {
        return this._properties[index];
      }

      var prop = new ICAL.Property(
        this.jCal[PROPERTY_INDEX][index],
        this
      );

      this._hydratedPropertyCount++;
      return (this._properties[index] = prop);
    },

    /**
     * Finds first sub component, optionally filtered by name.
     *
     * @param {String=} name        Optional name to filter by
     * @return {?ICAL.Component}     The found subcomponent
     */
    getFirstSubcomponent: function(name) {
      if (name) {
        var i = 0;
        var comps = this.jCal[COMPONENT_INDEX];
        var len = comps.length;

        for (; i < len; i++) {
          if (comps[i][NAME_INDEX] === name) {
            var result = this._hydrateComponent(i);
            return result;
          }
        }
      } else {
        if (this.jCal[COMPONENT_INDEX].length) {
          return this._hydrateComponent(0);
        }
      }

      // ensure we return a value (strict mode)
      return null;
    },

    /**
     * Finds all sub components, optionally filtering by name.
     *
     * @param {String=} name            Optional name to filter by
     * @return {ICAL.Component[]}       The found sub components
     */
    getAllSubcomponents: function(name) {
      var jCalLen = this.jCal[COMPONENT_INDEX].length;
      var i = 0;

      if (name) {
        var comps = this.jCal[COMPONENT_INDEX];
        var result = [];

        for (; i < jCalLen; i++) {
          if (name === comps[i][NAME_INDEX]) {
            result.push(
              this._hydrateComponent(i)
            );
          }
        }
        return result;
      } else {
        if (!this._components ||
            (this._hydratedComponentCount !== jCalLen)) {
          for (; i < jCalLen; i++) {
            this._hydrateComponent(i);
          }
        }

        return this._components || [];
      }
    },

    /**
     * Returns true when a named property exists.
     *
     * @param {String} name     The property name
     * @return {Boolean}        True, when property is found
     */
    hasProperty: function(name) {
      var props = this.jCal[PROPERTY_INDEX];
      var len = props.length;

      var i = 0;
      for (; i < len; i++) {
        // 0 is property name
        if (props[i][NAME_INDEX] === name) {
          return true;
        }
      }

      return false;
    },

    /**
     * Finds the first property, optionally with the given name.
     *
     * @param {String=} name        Lowercase property name
     * @return {?ICAL.Property}     The found property
     */
    getFirstProperty: function(name) {
      if (name) {
        var i = 0;
        var props = this.jCal[PROPERTY_INDEX];
        var len = props.length;

        for (; i < len; i++) {
          if (props[i][NAME_INDEX] === name) {
            var result = this._hydrateProperty(i);
            return result;
          }
        }
      } else {
        if (this.jCal[PROPERTY_INDEX].length) {
          return this._hydrateProperty(0);
        }
      }

      return null;
    },

    /**
     * Returns first property's value, if available.
     *
     * @param {String=} name    Lowercase property name
     * @return {?String}        The found property value.
     */
    getFirstPropertyValue: function(name) {
      var prop = this.getFirstProperty(name);
      if (prop) {
        return prop.getFirstValue();
      }

      return null;
    },

    /**
     * Get all properties in the component, optionally filtered by name.
     *
     * @param {String=} name        Lowercase property name
     * @return {ICAL.Property[]}    List of properties
     */
    getAllProperties: function(name) {
      var jCalLen = this.jCal[PROPERTY_INDEX].length;
      var i = 0;

      if (name) {
        var props = this.jCal[PROPERTY_INDEX];
        var result = [];

        for (; i < jCalLen; i++) {
          if (name === props[i][NAME_INDEX]) {
            result.push(
              this._hydrateProperty(i)
            );
          }
        }
        return result;
      } else {
        if (!this._properties ||
            (this._hydratedPropertyCount !== jCalLen)) {
          for (; i < jCalLen; i++) {
            this._hydrateProperty(i);
          }
        }

        return this._properties || [];
      }
    },

    _removeObjectByIndex: function(jCalIndex, cache, index) {
      cache = cache || [];
      // remove cached version
      if (cache[index]) {
        var obj = cache[index];
        if ("parent" in obj) {
            obj.parent = null;
        }
      }

      cache.splice(index, 1);

      // remove it from the jCal
      this.jCal[jCalIndex].splice(index, 1);
    },

    _removeObject: function(jCalIndex, cache, nameOrObject) {
      var i = 0;
      var objects = this.jCal[jCalIndex];
      var len = objects.length;
      var cached = this[cache];

      if (typeof(nameOrObject) === 'string') {
        for (; i < len; i++) {
          if (objects[i][NAME_INDEX] === nameOrObject) {
            this._removeObjectByIndex(jCalIndex, cached, i);
            return true;
          }
        }
      } else if (cached) {
        for (; i < len; i++) {
          if (cached[i] && cached[i] === nameOrObject) {
            this._removeObjectByIndex(jCalIndex, cached, i);
            return true;
          }
        }
      }

      return false;
    },

    _removeAllObjects: function(jCalIndex, cache, name) {
      var cached = this[cache];

      // Unfortunately we have to run through all children to reset their
      // parent property.
      var objects = this.jCal[jCalIndex];
      var i = objects.length - 1;

      // descending search required because splice
      // is used and will effect the indices.
      for (; i >= 0; i--) {
        if (!name || objects[i][NAME_INDEX] === name) {
          this._removeObjectByIndex(jCalIndex, cached, i);
        }
      }
    },

    /**
     * Adds a single sub component.
     *
     * @param {ICAL.Component} component        The component to add
     * @return {ICAL.Component}                 The passed in component
     */
    addSubcomponent: function(component) {
      if (!this._components) {
        this._components = [];
        this._hydratedComponentCount = 0;
      }

      if (component.parent) {
        component.parent.removeSubcomponent(component);
      }

      var idx = this.jCal[COMPONENT_INDEX].push(component.jCal);
      this._components[idx - 1] = component;
      this._hydratedComponentCount++;
      component.parent = this;
      return component;
    },

    /**
     * Removes a single component by name or the instance of a specific
     * component.
     *
     * @param {ICAL.Component|String} nameOrComp    Name of component, or component
     * @return {Boolean}                            True when comp is removed
     */
    removeSubcomponent: function(nameOrComp) {
      var removed = this._removeObject(COMPONENT_INDEX, '_components', nameOrComp);
      if (removed) {
        this._hydratedComponentCount--;
      }
      return removed;
    },

    /**
     * Removes all components or (if given) all components by a particular
     * name.
     *
     * @param {String=} name            Lowercase component name
     */
    removeAllSubcomponents: function(name) {
      var removed = this._removeAllObjects(COMPONENT_INDEX, '_components', name);
      this._hydratedComponentCount = 0;
      return removed;
    },

    /**
     * Adds an {@link ICAL.Property} to the component.
     *
     * @param {ICAL.Property} property      The property to add
     * @return {ICAL.Property}              The passed in property
     */
    addProperty: function(property) {
      if (!(property instanceof ICAL.Property)) {
        throw new TypeError('must instance of ICAL.Property');
      }

      if (!this._properties) {
        this._properties = [];
        this._hydratedPropertyCount = 0;
      }

      if (property.parent) {
        property.parent.removeProperty(property);
      }

      var idx = this.jCal[PROPERTY_INDEX].push(property.jCal);
      this._properties[idx - 1] = property;
      this._hydratedPropertyCount++;
      property.parent = this;
      return property;
    },

    /**
     * Helper method to add a property with a value to the component.
     *
     * @param {String}               name         Property name to add
     * @param {String|Number|Object} value        Property value
     * @return {ICAL.Property}                    The created property
     */
    addPropertyWithValue: function(name, value) {
      var prop = new ICAL.Property(name);
      prop.setValue(value);

      this.addProperty(prop);

      return prop;
    },

    /**
     * Helper method that will update or create a property of the given name
     * and sets its value. If multiple properties with the given name exist,
     * only the first is updated.
     *
     * @param {String}               name         Property name to update
     * @param {String|Number|Object} value        Property value
     * @return {ICAL.Property}                    The created property
     */
    updatePropertyWithValue: function(name, value) {
      var prop = this.getFirstProperty(name);

      if (prop) {
        prop.setValue(value);
      } else {
        prop = this.addPropertyWithValue(name, value);
      }

      return prop;
    },

    /**
     * Removes a single property by name or the instance of the specific
     * property.
     *
     * @param {String|ICAL.Property} nameOrProp     Property name or instance to remove
     * @return {Boolean}                            True, when deleted
     */
    removeProperty: function(nameOrProp) {
      var removed = this._removeObject(PROPERTY_INDEX, '_properties', nameOrProp);
      if (removed) {
        this._hydratedPropertyCount--;
      }
      return removed;
    },

    /**
     * Removes all properties associated with this component, optionally
     * filtered by name.
     *
     * @param {String=} name        Lowercase property name
     * @return {Boolean}            True, when deleted
     */
    removeAllProperties: function(name) {
      var removed = this._removeAllObjects(PROPERTY_INDEX, '_properties', name);
      this._hydratedPropertyCount = 0;
      return removed;
    },

    /**
     * Returns the Object representation of this component. The returned object
     * is a live jCal object and should be cloned if modified.
     * @return {Object}
     */
    toJSON: function() {
      return this.jCal;
    },

    /**
     * The string representation of this component.
     * @return {String}
     */
    toString: function() {
      return ICAL.stringify.component(
        this.jCal, this._designSet
      );
    }
  };

  /**
   * Create an {@link ICAL.Component} by parsing the passed iCalendar string.
   *
   * @param {String} str        The iCalendar string to parse
   */
  Component.fromString = function(str) {
    return new Component(ICAL.parse.component(str));
  };

  return Component;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.Property = (function() {
  'use strict';

  var NAME_INDEX = 0;
  var PROP_INDEX = 1;
  var TYPE_INDEX = 2;
  var VALUE_INDEX = 3;

  var design = ICAL.design;

  /**
   * @classdesc
   * Provides a layer on top of the raw jCal object for manipulating a single
   * property, with its parameters and value.
   *
   * @description
   * Its important to note that mutations done in the wrapper
   * directly mutate the jCal object used to initialize.
   *
   * Can also be used to create new properties by passing
   * the name of the property (as a String).
   *
   * @class
   * @alias ICAL.Property
   * @param {Array|String} jCal         Raw jCal representation OR
   *  the new name of the property
   *
   * @param {ICAL.Component=} parent    Parent component
   */
  function Property(jCal, parent) {
    this._parent = parent || null;

    if (typeof(jCal) === 'string') {
      // We are creating the property by name and need to detect the type
      this.jCal = [jCal, {}, design.defaultType];
      this.jCal[TYPE_INDEX] = this.getDefaultType();
    } else {
      this.jCal = jCal;
    }
    this._updateType();
  }

  Property.prototype = {

    /**
     * The value type for this property
     * @readonly
     * @type {String}
     */
    get type() {
      return this.jCal[TYPE_INDEX];
    },

    /**
     * The name of this property, in lowercase.
     * @readonly
     * @type {String}
     */
    get name() {
      return this.jCal[NAME_INDEX];
    },

    /**
     * The parent component for this property.
     * @type {ICAL.Component}
     */
    get parent() {
      return this._parent;
    },

    set parent(p) {
      // Before setting the parent, check if the design set has changed. If it
      // has, we later need to update the type if it was unknown before.
      var designSetChanged = !this._parent || (p && p._designSet != this._parent._designSet);

      this._parent = p;

      if (this.type == design.defaultType && designSetChanged) {
        this.jCal[TYPE_INDEX] = this.getDefaultType();
        this._updateType();
      }

      return p;
    },

    /**
     * The design set for this property, e.g. icalendar vs vcard
     *
     * @type {ICAL.design.designSet}
     * @private
     */
    get _designSet() {
      return this.parent ? this.parent._designSet : design.defaultSet;
    },

    /**
     * Updates the type metadata from the current jCal type and design set.
     *
     * @private
     */
    _updateType: function() {
      var designSet = this._designSet;

      if (this.type in designSet.value) {
        var designType = designSet.value[this.type];

        if ('decorate' in designSet.value[this.type]) {
          this.isDecorated = true;
        } else {
          this.isDecorated = false;
        }

        if (this.name in designSet.property) {
          this.isMultiValue = ('multiValue' in designSet.property[this.name]);
          this.isStructuredValue = ('structuredValue' in designSet.property[this.name]);
        }
      }
    },

    /**
     * Hydrate a single value. The act of hydrating means turning the raw jCal
     * value into a potentially wrapped object, for example {@link ICAL.Time}.
     *
     * @private
     * @param {Number} index        The index of the value to hydrate
     * @return {Object}             The decorated value.
     */
    _hydrateValue: function(index) {
      if (this._values && this._values[index]) {
        return this._values[index];
      }

      // for the case where there is no value.
      if (this.jCal.length <= (VALUE_INDEX + index)) {
        return null;
      }

      if (this.isDecorated) {
        if (!this._values) {
          this._values = [];
        }
        return (this._values[index] = this._decorate(
          this.jCal[VALUE_INDEX + index]
        ));
      } else {
        return this.jCal[VALUE_INDEX + index];
      }
    },

    /**
     * Decorate a single value, returning its wrapped object. This is used by
     * the hydrate function to actually wrap the value.
     *
     * @private
     * @param {?} value         The value to decorate
     * @return {Object}         The decorated value
     */
    _decorate: function(value) {
      return this._designSet.value[this.type].decorate(value, this);
    },

    /**
     * Undecorate a single value, returning its raw jCal data.
     *
     * @private
     * @param {Object} value         The value to undecorate
     * @return {?}                   The undecorated value
     */
    _undecorate: function(value) {
      return this._designSet.value[this.type].undecorate(value, this);
    },

    /**
     * Sets the value at the given index while also hydrating it. The passed
     * value can either be a decorated or undecorated value.
     *
     * @private
     * @param {?} value             The value to set
     * @param {Number} index        The index to set it at
     */
    _setDecoratedValue: function(value, index) {
      if (!this._values) {
        this._values = [];
      }

      if (typeof(value) === 'object' && 'icaltype' in value) {
        // decorated value
        this.jCal[VALUE_INDEX + index] = this._undecorate(value);
        this._values[index] = value;
      } else {
        // undecorated value
        this.jCal[VALUE_INDEX + index] = value;
        this._values[index] = this._decorate(value);
      }
    },

    /**
     * Gets a parameter on the property.
     *
     * @param {String}        name   Property name (lowercase)
     * @return {Array|String}        Property value
     */
    getParameter: function(name) {
      if (name in this.jCal[PROP_INDEX]) {
        return this.jCal[PROP_INDEX][name];
      } else {
        return undefined;
      }
    },

    /**
     * Gets first parameter on the property.
     *
     * @param {String}        name   Property name (lowercase)
     * @return {String}        Property value
     */
    getFirstParameter: function(name) {
      var parameters = this.getParameter(name);

      if (Array.isArray(parameters)) {
        return parameters[0];
      }

      return parameters;
    },

    /**
     * Sets a parameter on the property.
     *
     * @param {String}       name     The parameter name
     * @param {Array|String} value    The parameter value
     */
    setParameter: function(name, value) {
      var lcname = name.toLowerCase();
      if (typeof value === "string" &&
          lcname in this._designSet.param &&
          'multiValue' in this._designSet.param[lcname]) {
          value = [value];
      }
      this.jCal[PROP_INDEX][name] = value;
    },

    /**
     * Removes a parameter
     *
     * @param {String} name     The parameter name
     */
    removeParameter: function(name) {
      delete this.jCal[PROP_INDEX][name];
    },

    /**
     * Get the default type based on this property's name.
     *
     * @return {String}     The default type for this property
     */
    getDefaultType: function() {
      var name = this.jCal[NAME_INDEX];
      var designSet = this._designSet;

      if (name in designSet.property) {
        var details = designSet.property[name];
        if ('defaultType' in details) {
          return details.defaultType;
        }
      }
      return design.defaultType;
    },

    /**
     * Sets type of property and clears out any existing values of the current
     * type.
     *
     * @param {String} type     New iCAL type (see design.*.values)
     */
    resetType: function(type) {
      this.removeAllValues();
      this.jCal[TYPE_INDEX] = type;
      this._updateType();
    },

    /**
     * Finds the first property value.
     *
     * @return {String}         First property value
     */
    getFirstValue: function() {
      return this._hydrateValue(0);
    },

    /**
     * Gets all values on the property.
     *
     * NOTE: this creates an array during each call.
     *
     * @return {Array}          List of values
     */
    getValues: function() {
      var len = this.jCal.length - VALUE_INDEX;

      if (len < 1) {
        // its possible for a property to have no value.
        return [];
      }

      var i = 0;
      var result = [];

      for (; i < len; i++) {
        result[i] = this._hydrateValue(i);
      }

      return result;
    },

    /**
     * Removes all values from this property
     */
    removeAllValues: function() {
      if (this._values) {
        this._values.length = 0;
      }
      this.jCal.length = 3;
    },

    /**
     * Sets the values of the property.  Will overwrite the existing values.
     * This can only be used for multi-value properties.
     *
     * @param {Array} values    An array of values
     */
    setValues: function(values) {
      if (!this.isMultiValue) {
        throw new Error(
          this.name + ': does not not support mulitValue.\n' +
          'override isMultiValue'
        );
      }

      var len = values.length;
      var i = 0;
      this.removeAllValues();

      if (len > 0 &&
          typeof(values[0]) === 'object' &&
          'icaltype' in values[0]) {
        this.resetType(values[0].icaltype);
      }

      if (this.isDecorated) {
        for (; i < len; i++) {
          this._setDecoratedValue(values[i], i);
        }
      } else {
        for (; i < len; i++) {
          this.jCal[VALUE_INDEX + i] = values[i];
        }
      }
    },

    /**
     * Sets the current value of the property. If this is a multi-value
     * property, all other values will be removed.
     *
     * @param {String|Object} value     New property value.
     */
    setValue: function(value) {
      this.removeAllValues();
      if (typeof(value) === 'object' && 'icaltype' in value) {
        this.resetType(value.icaltype);
      }

      if (this.isDecorated) {
        this._setDecoratedValue(value, 0);
      } else {
        this.jCal[VALUE_INDEX] = value;
      }
    },

    /**
     * Returns the Object representation of this component. The returned object
     * is a live jCal object and should be cloned if modified.
     * @return {Object}
     */
    toJSON: function() {
      return this.jCal;
    },

    /**
     * The string representation of this component.
     * @return {String}
     */
    toICALString: function() {
      return ICAL.stringify.property(
        this.jCal, this._designSet, true
      );
    }
  };

  /**
   * Create an {@link ICAL.Property} by parsing the passed iCalendar string.
   *
   * @param {String} str                        The iCalendar string to parse
   * @param {ICAL.design.designSet=} designSet  The design data to use for this property
   * @return {ICAL.Property}                    The created iCalendar property
   */
  Property.fromString = function(str, designSet) {
    return new Property(ICAL.parse.property(str, designSet));
  };

  return Property;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.UtcOffset = (function() {

  /**
   * @classdesc
   * This class represents the "duration" value type, with various calculation
   * and manipulation methods.
   *
   * @class
   * @alias ICAL.UtcOffset
   * @param {Object} aData          An object with members of the utc offset
   * @param {Number=} aData.hours   The hours for the utc offset
   * @param {Number=} aData.minutes The minutes in the utc offset
   * @param {Number=} aData.factor  The factor for the utc-offset, either -1 or 1
   */
  function UtcOffset(aData) {
    this.fromData(aData);
  }

  UtcOffset.prototype = {

    /**
     * The hours in the utc-offset
     * @type {Number}
     */
    hours: 0,

    /**
     * The minutes in the utc-offset
     * @type {Number}
     */
    minutes: 0,

    /**
     * The sign of the utc offset, 1 for positive offset, -1 for negative
     * offsets.
     * @type {Number}
     */
    factor: 1,

    /**
     * The type name, to be used in the jCal object.
     * @constant
     * @type {String}
     * @default "utc-offset"
     */
    icaltype: "utc-offset",

    /**
     * Returns a clone of the utc offset object.
     *
     * @return {ICAL.UtcOffset}     The cloned object
     */
    clone: function() {
      return ICAL.UtcOffset.fromSeconds(this.toSeconds());
    },

    /**
     * Sets up the current instance using members from the passed data object.
     *
     * @param {Object} aData          An object with members of the utc offset
     * @param {Number=} aData.hours   The hours for the utc offset
     * @param {Number=} aData.minutes The minutes in the utc offset
     * @param {Number=} aData.factor  The factor for the utc-offset, either -1 or 1
     */
    fromData: function(aData) {
      if (aData) {
        for (var key in aData) {
          /* istanbul ignore else */
          if (aData.hasOwnProperty(key)) {
            this[key] = aData[key];
          }
        }
      }
      this._normalize();
    },

    /**
     * Sets up the current instance from the given seconds value. The seconds
     * value is truncated to the minute. Offsets are wrapped when the world
     * ends, the hour after UTC+14:00 is UTC-12:00.
     *
     * @param {Number} aSeconds         The seconds to convert into an offset
     */
    fromSeconds: function(aSeconds) {
      var secs = Math.abs(aSeconds);

      this.factor = aSeconds < 0 ? -1 : 1;
      this.hours = ICAL.helpers.trunc(secs / 3600);

      secs -= (this.hours * 3600);
      this.minutes = ICAL.helpers.trunc(secs / 60);
      return this;
    },

    /**
     * Convert the current offset to a value in seconds
     *
     * @return {Number}                 The offset in seconds
     */
    toSeconds: function() {
      return this.factor * (60 * this.minutes + 3600 * this.hours);
    },

    /**
     * Compare this utc offset with another one.
     *
     * @param {ICAL.UtcOffset} other        The other offset to compare with
     * @return {Number}                     -1, 0 or 1 for less/equal/greater
     */
    compare: function icaltime_compare(other) {
      var a = this.toSeconds();
      var b = other.toSeconds();
      return (a > b) - (b > a);
    },

    _normalize: function() {
      // Range: 97200 seconds (with 1 hour inbetween)
      var secs = this.toSeconds();
      var factor = this.factor;
      while (secs < -43200) { // = UTC-12:00
        secs += 97200;
      }
      while (secs > 50400) { // = UTC+14:00
        secs -= 97200;
      }

      this.fromSeconds(secs);

      // Avoid changing the factor when on zero seconds
      if (secs == 0) {
        this.factor = factor;
      }
    },

    /**
     * The iCalendar string representation of this utc-offset.
     * @return {String}
     */
    toICALString: function() {
      return ICAL.design.icalendar.value['utc-offset'].toICAL(this.toString());
    },

    /**
     * The string representation of this utc-offset.
     * @return {String}
     */
    toString: function toString() {
      return (this.factor == 1 ? "+" : "-") +
              ICAL.helpers.pad2(this.hours) + ':' +
              ICAL.helpers.pad2(this.minutes);
    }
  };

  /**
   * Creates a new {@link ICAL.UtcOffset} instance from the passed string.
   *
   * @param {String} aString    The string to parse
   * @return {ICAL.Duration}    The created utc-offset instance
   */
  UtcOffset.fromString = function(aString) {
    // -05:00
    var options = {};
    //TODO: support seconds per rfc5545 ?
    options.factor = (aString[0] === '+') ? 1 : -1;
    options.hours = ICAL.helpers.strictParseInt(aString.substr(1, 2));
    options.minutes = ICAL.helpers.strictParseInt(aString.substr(4, 2));

    return new ICAL.UtcOffset(options);
  };

  /**
   * Creates a new {@link ICAL.UtcOffset} instance from the passed seconds
   * value.
   *
   * @param {Number} aSeconds       The number of seconds to convert
   */
  UtcOffset.fromSeconds = function(aSeconds) {
    var instance = new UtcOffset();
    instance.fromSeconds(aSeconds);
    return instance;
  };

  return UtcOffset;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.Binary = (function() {

  /**
   * @classdesc
   * Represents the BINARY value type, which contains extra methods for
   * encoding and decoding.
   *
   * @class
   * @alias ICAL.Binary
   * @param {String} aValue     The binary data for this value
   */
  function Binary(aValue) {
    this.value = aValue;
  }

  Binary.prototype = {
    /**
     * The type name, to be used in the jCal object.
     * @default "binary"
     * @constant
     */
    icaltype: "binary",

    /**
     * Base64 decode the current value
     *
     * @return {String}         The base64-decoded value
     */
    decodeValue: function decodeValue() {
      return this._b64_decode(this.value);
    },

    /**
     * Encodes the passed parameter with base64 and sets the internal
     * value to the result.
     *
     * @param {String} aValue      The raw binary value to encode
     */
    setEncodedValue: function setEncodedValue(aValue) {
      this.value = this._b64_encode(aValue);
    },

    _b64_encode: function base64_encode(data) {
      // http://kevin.vanzonneveld.net
      // +   original by: Tyler Akins (http://rumkin.com)
      // +   improved by: Bayron Guevara
      // +   improved by: Thunder.m
      // +   improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
      // +   bugfixed by: Pellentesque Malesuada
      // +   improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
      // +   improved by: Rafał Kukawski (http://kukawski.pl)
      // *     example 1: base64_encode('Kevin van Zonneveld');
      // *     returns 1: 'S2V2aW4gdmFuIFpvbm5ldmVsZA=='
      // mozilla has this native
      // - but breaks in 2.0.0.12!
      //if (typeof this.window['atob'] == 'function') {
      //    return atob(data);
      //}
      var b64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" +
                "abcdefghijklmnopqrstuvwxyz0123456789+/=";
      var o1, o2, o3, h1, h2, h3, h4, bits, i = 0,
        ac = 0,
        enc = "",
        tmp_arr = [];

      if (!data) {
        return data;
      }

      do { // pack three octets into four hexets
        o1 = data.charCodeAt(i++);
        o2 = data.charCodeAt(i++);
        o3 = data.charCodeAt(i++);

        bits = o1 << 16 | o2 << 8 | o3;

        h1 = bits >> 18 & 0x3f;
        h2 = bits >> 12 & 0x3f;
        h3 = bits >> 6 & 0x3f;
        h4 = bits & 0x3f;

        // use hexets to index into b64, and append result to encoded string
        tmp_arr[ac++] = b64.charAt(h1) + b64.charAt(h2) + b64.charAt(h3) + b64.charAt(h4);
      } while (i < data.length);

      enc = tmp_arr.join('');

      var r = data.length % 3;

      return (r ? enc.slice(0, r - 3) : enc) + '==='.slice(r || 3);

    },

    _b64_decode: function base64_decode(data) {
      // http://kevin.vanzonneveld.net
      // +   original by: Tyler Akins (http://rumkin.com)
      // +   improved by: Thunder.m
      // +      input by: Aman Gupta
      // +   improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
      // +   bugfixed by: Onno Marsman
      // +   bugfixed by: Pellentesque Malesuada
      // +   improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
      // +      input by: Brett Zamir (http://brett-zamir.me)
      // +   bugfixed by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
      // *     example 1: base64_decode('S2V2aW4gdmFuIFpvbm5ldmVsZA==');
      // *     returns 1: 'Kevin van Zonneveld'
      // mozilla has this native
      // - but breaks in 2.0.0.12!
      //if (typeof this.window['btoa'] == 'function') {
      //    return btoa(data);
      //}
      var b64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" +
                "abcdefghijklmnopqrstuvwxyz0123456789+/=";
      var o1, o2, o3, h1, h2, h3, h4, bits, i = 0,
        ac = 0,
        dec = "",
        tmp_arr = [];

      if (!data) {
        return data;
      }

      data += '';

      do { // unpack four hexets into three octets using index points in b64
        h1 = b64.indexOf(data.charAt(i++));
        h2 = b64.indexOf(data.charAt(i++));
        h3 = b64.indexOf(data.charAt(i++));
        h4 = b64.indexOf(data.charAt(i++));

        bits = h1 << 18 | h2 << 12 | h3 << 6 | h4;

        o1 = bits >> 16 & 0xff;
        o2 = bits >> 8 & 0xff;
        o3 = bits & 0xff;

        if (h3 == 64) {
          tmp_arr[ac++] = String.fromCharCode(o1);
        } else if (h4 == 64) {
          tmp_arr[ac++] = String.fromCharCode(o1, o2);
        } else {
          tmp_arr[ac++] = String.fromCharCode(o1, o2, o3);
        }
      } while (i < data.length);

      dec = tmp_arr.join('');

      return dec;
    },

    /**
     * The string representation of this value
     * @return {String}
     */
    toString: function() {
      return this.value;
    }
  };

  /**
   * Creates a binary value from the given string.
   *
   * @param {String} aString        The binary value string
   * @return {ICAL.Binary}          The binary value instance
   */
  Binary.fromString = function(aString) {
    return new Binary(aString);
  };

  return Binary;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */



(function() {
  /**
   * @classdesc
   * This class represents the "period" value type, with various calculation
   * and manipulation methods.
   *
   * @description
   * The passed data object cannot contain both and end date and a duration.
   *
   * @class
   * @param {Object} aData                  An object with members of the period
   * @param {ICAL.Time=} aData.start        The start of the period
   * @param {ICAL.Time=} aData.end          The end of the period
   * @param {ICAL.Duration=} aData.duration The duration of the period
   */
  ICAL.Period = function icalperiod(aData) {
    this.wrappedJSObject = this;

    if (aData && 'start' in aData) {
      if (aData.start && !(aData.start instanceof ICAL.Time)) {
        throw new TypeError('.start must be an instance of ICAL.Time');
      }
      this.start = aData.start;
    }

    if (aData && aData.end && aData.duration) {
      throw new Error('cannot accept both end and duration');
    }

    if (aData && 'end' in aData) {
      if (aData.end && !(aData.end instanceof ICAL.Time)) {
        throw new TypeError('.end must be an instance of ICAL.Time');
      }
      this.end = aData.end;
    }

    if (aData && 'duration' in aData) {
      if (aData.duration && !(aData.duration instanceof ICAL.Duration)) {
        throw new TypeError('.duration must be an instance of ICAL.Duration');
      }
      this.duration = aData.duration;
    }
  };

  ICAL.Period.prototype = {

    /**
     * The start of the period
     * @type {ICAL.Time}
     */
    start: null,

    /**
     * The end of the period
     * @type {ICAL.Time}
     */
    end: null,

    /**
     * The duration of the period
     * @type {ICAL.Duration}
     */
    duration: null,

    /**
     * The class identifier.
     * @constant
     * @type {String}
     * @default "icalperiod"
     */
    icalclass: "icalperiod",

    /**
     * The type name, to be used in the jCal object.
     * @constant
     * @type {String}
     * @default "period"
     */
    icaltype: "period",

    /**
     * Returns a clone of the duration object.
     *
     * @return {ICAL.Period}      The cloned object
     */
    clone: function() {
      return ICAL.Period.fromData({
        start: this.start ? this.start.clone() : null,
        end: this.end ? this.end.clone() : null,
        duration: this.duration ? this.duration.clone() : null
      });
    },

    /**
     * Calculates the duration of the period, either directly or by subtracting
     * start from end date.
     *
     * @return {ICAL.Duration}      The calculated duration
     */
    getDuration: function duration() {
      if (this.duration) {
        return this.duration;
      } else {
        return this.end.subtractDate(this.start);
      }
    },

    /**
     * Calculates the end date of the period, either directly or by adding
     * duration to start date.
     *
     * @return {ICAL.Time}          The calculated end date
     */
    getEnd: function() {
      if (this.end) {
        return this.end;
      } else {
        var end = this.start.clone();
        end.addDuration(this.duration);
        return end;
      }
    },

    /**
     * The string representation of this period.
     * @return {String}
     */
    toString: function toString() {
      return this.start + "/" + (this.end || this.duration);
    },

    /**
     * The jCal representation of this period type.
     * @return {Object}
     */
    toJSON: function() {
      return [this.start.toString(), (this.end || this.duration).toString()];
    },

    /**
     * The iCalendar string representation of this period.
     * @return {String}
     */
    toICALString: function() {
      return this.start.toICALString() + "/" +
             (this.end || this.duration).toICALString();
    }
  };

  /**
   * Creates a new {@link ICAL.Period} instance from the passed string.
   *
   * @param {String} str            The string to parse
   * @param {ICAL.Property} prop    The property this period will be on
   * @return {ICAL.Period}          The created period instance
   */
  ICAL.Period.fromString = function fromString(str, prop) {
    var parts = str.split('/');

    if (parts.length !== 2) {
      throw new Error(
        'Invalid string value: "' + str + '" must contain a "/" char.'
      );
    }

    var options = {
      start: ICAL.Time.fromDateTimeString(parts[0], prop)
    };

    var end = parts[1];

    if (ICAL.Duration.isValueString(end)) {
      options.duration = ICAL.Duration.fromString(end);
    } else {
      options.end = ICAL.Time.fromDateTimeString(end, prop);
    }

    return new ICAL.Period(options);
  };

  /**
   * Creates a new {@link ICAL.Period} instance from the given data object.
   * The passed data object cannot contain both and end date and a duration.
   *
   * @param {Object} aData                  An object with members of the period
   * @param {ICAL.Time=} aData.start        The start of the period
   * @param {ICAL.Time=} aData.end          The end of the period
   * @param {ICAL.Duration=} aData.duration The duration of the period
   * @return {ICAL.Period}                  The period instance
   */
  ICAL.Period.fromData = function fromData(aData) {
    return new ICAL.Period(aData);
  };

  /**
   * Returns a new period instance from the given jCal data array. The first
   * member is always the start date string, the second member is either a
   * duration or end date string.
   *
   * @param {Array<String,String>} aData    The jCal data array
   * @param {ICAL.Property} aProp           The property this jCal data is on
   * @param {Boolean} aLenient              If true, data value can be both date and date-time
   * @return {ICAL.Period}                  The period instance
   */
  ICAL.Period.fromJSON = function(aData, aProp, aLenient) {
    function fromDateOrDateTimeString(aValue, aProp) {
      if (aLenient) {
        return ICAL.Time.fromString(aValue, aProp);
      } else {
        return ICAL.Time.fromDateTimeString(aValue, aProp);
      }
    }

    if (ICAL.Duration.isValueString(aData[1])) {
      return ICAL.Period.fromData({
        start: fromDateOrDateTimeString(aData[0], aProp),
        duration: ICAL.Duration.fromString(aData[1])
      });
    } else {
      return ICAL.Period.fromData({
        start: fromDateOrDateTimeString(aData[0], aProp),
        end: fromDateOrDateTimeString(aData[1], aProp)
      });
    }
  };
})();
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */



(function() {
  var DURATION_LETTERS = /([PDWHMTS]{1,1})/;

  /**
   * @classdesc
   * This class represents the "duration" value type, with various calculation
   * and manipulation methods.
   *
   * @class
   * @alias ICAL.Duration
   * @param {Object} data               An object with members of the duration
   * @param {Number} data.weeks         Duration in weeks
   * @param {Number} data.days          Duration in days
   * @param {Number} data.hours         Duration in hours
   * @param {Number} data.minutes       Duration in minutes
   * @param {Number} data.seconds       Duration in seconds
   * @param {Boolean} data.isNegative   If true, the duration is negative
   */
  ICAL.Duration = function icalduration(data) {
    this.wrappedJSObject = this;
    this.fromData(data);
  };

  ICAL.Duration.prototype = {
    /**
     * The weeks in this duration
     * @type {Number}
     * @default 0
     */
    weeks: 0,

    /**
     * The days in this duration
     * @type {Number}
     * @default 0
     */
    days: 0,

    /**
     * The days in this duration
     * @type {Number}
     * @default 0
     */
    hours: 0,

    /**
     * The minutes in this duration
     * @type {Number}
     * @default 0
     */
    minutes: 0,

    /**
     * The seconds in this duration
     * @type {Number}
     * @default 0
     */
    seconds: 0,

    /**
     * The seconds in this duration
     * @type {Boolean}
     * @default false
     */
    isNegative: false,

    /**
     * The class identifier.
     * @constant
     * @type {String}
     * @default "icalduration"
     */
    icalclass: "icalduration",

    /**
     * The type name, to be used in the jCal object.
     * @constant
     * @type {String}
     * @default "duration"
     */
    icaltype: "duration",

    /**
     * Returns a clone of the duration object.
     *
     * @return {ICAL.Duration}      The cloned object
     */
    clone: function clone() {
      return ICAL.Duration.fromData(this);
    },

    /**
     * The duration value expressed as a number of seconds.
     *
     * @return {Number}             The duration value in seconds
     */
    toSeconds: function toSeconds() {
      var seconds = this.seconds + 60 * this.minutes + 3600 * this.hours +
                    86400 * this.days + 7 * 86400 * this.weeks;
      return (this.isNegative ? -seconds : seconds);
    },

    /**
     * Reads the passed seconds value into this duration object. Afterwards,
     * members like {@link ICAL.Duration#days days} and {@link ICAL.Duration#weeks weeks} will be set up
     * accordingly.
     *
     * @param {Number} aSeconds     The duration value in seconds
     * @return {ICAL.Duration}      Returns this instance
     */
    fromSeconds: function fromSeconds(aSeconds) {
      var secs = Math.abs(aSeconds);

      this.isNegative = (aSeconds < 0);
      this.days = ICAL.helpers.trunc(secs / 86400);

      // If we have a flat number of weeks, use them.
      if (this.days % 7 == 0) {
        this.weeks = this.days / 7;
        this.days = 0;
      } else {
        this.weeks = 0;
      }

      secs -= (this.days + 7 * this.weeks) * 86400;

      this.hours = ICAL.helpers.trunc(secs / 3600);
      secs -= this.hours * 3600;

      this.minutes = ICAL.helpers.trunc(secs / 60);
      secs -= this.minutes * 60;

      this.seconds = secs;
      return this;
    },

    /**
     * Sets up the current instance using members from the passed data object.
     *
     * @param {Object} aData               An object with members of the duration
     * @param {Number} aData.weeks         Duration in weeks
     * @param {Number} aData.days          Duration in days
     * @param {Number} aData.hours         Duration in hours
     * @param {Number} aData.minutes       Duration in minutes
     * @param {Number} aData.seconds       Duration in seconds
     * @param {Boolean} aData.isNegative   If true, the duration is negative
     */
    fromData: function fromData(aData) {
      var propsToCopy = ["weeks", "days", "hours",
                         "minutes", "seconds", "isNegative"];
      for (var key in propsToCopy) {
        /* istanbul ignore if */
        if (!propsToCopy.hasOwnProperty(key)) {
          continue;
        }
        var prop = propsToCopy[key];
        if (aData && prop in aData) {
          this[prop] = aData[prop];
        } else {
          this[prop] = 0;
        }
      }
    },

    /**
     * Resets the duration instance to the default values, i.e. PT0S
     */
    reset: function reset() {
      this.isNegative = false;
      this.weeks = 0;
      this.days = 0;
      this.hours = 0;
      this.minutes = 0;
      this.seconds = 0;
    },

    /**
     * Compares the duration instance with another one.
     *
     * @param {ICAL.Duration} aOther        The instance to compare with
     * @return {Number}                     -1, 0 or 1 for less/equal/greater
     */
    compare: function compare(aOther) {
      var thisSeconds = this.toSeconds();
      var otherSeconds = aOther.toSeconds();
      return (thisSeconds > otherSeconds) - (thisSeconds < otherSeconds);
    },

    /**
     * Normalizes the duration instance. For example, a duration with a value
     * of 61 seconds will be normalized to 1 minute and 1 second.
     */
    normalize: function normalize() {
      this.fromSeconds(this.toSeconds());
    },

    /**
     * The string representation of this duration.
     * @return {String}
     */
    toString: function toString() {
      if (this.toSeconds() == 0) {
        return "PT0S";
      } else {
        var str = "";
        if (this.isNegative) str += "-";
        str += "P";
        if (this.weeks) str += this.weeks + "W";
        if (this.days) str += this.days + "D";

        if (this.hours || this.minutes || this.seconds) {
          str += "T";
          if (this.hours) str += this.hours + "H";
          if (this.minutes) str += this.minutes + "M";
          if (this.seconds) str += this.seconds + "S";
        }
        return str;
      }
    },

    /**
     * The iCalendar string representation of this duration.
     * @return {String}
     */
    toICALString: function() {
      return this.toString();
    }
  };

  /**
   * Returns a new ICAL.Duration instance from the passed seconds value.
   *
   * @param {Number} aSeconds       The seconds to create the instance from
   * @return {ICAL.Duration}        The newly created duration instance
   */
  ICAL.Duration.fromSeconds = function icalduration_from_seconds(aSeconds) {
    return (new ICAL.Duration()).fromSeconds(aSeconds);
  };

  /**
   * Internal helper function to handle a chunk of a duration.
   *
   * @param {String} letter type of duration chunk
   * @param {String} number numeric value or -/+
   * @param {Object} dict target to assign values to
   */
  function parseDurationChunk(letter, number, object) {
    var type;
    switch (letter) {
      case 'P':
        if (number && number === '-') {
          object.isNegative = true;
        } else {
          object.isNegative = false;
        }
        // period
        break;
      case 'D':
        type = 'days';
        break;
      case 'W':
        type = 'weeks';
        break;
      case 'H':
        type = 'hours';
        break;
      case 'M':
        type = 'minutes';
        break;
      case 'S':
        type = 'seconds';
        break;
      default:
        // Not a valid chunk
        return 0;
    }

    if (type) {
      if (!number && number !== 0) {
        throw new Error(
          'invalid duration value: Missing number before "' + letter + '"'
        );
      }
      var num = parseInt(number, 10);
      if (ICAL.helpers.isStrictlyNaN(num)) {
        throw new Error(
          'invalid duration value: Invalid number "' + number + '" before "' + letter + '"'
        );
      }
      object[type] = num;
    }

    return 1;
  }

  /**
   * Checks if the given string is an iCalendar duration value.
   *
   * @param {String} value      The raw ical value
   * @return {Boolean}          True, if the given value is of the
   *                              duration ical type
   */
  ICAL.Duration.isValueString = function(string) {
    return (string[0] === 'P' || string[1] === 'P');
  };

  /**
   * Creates a new {@link ICAL.Duration} instance from the passed string.
   *
   * @param {String} aStr       The string to parse
   * @return {ICAL.Duration}    The created duration instance
   */
  ICAL.Duration.fromString = function icalduration_from_string(aStr) {
    var pos = 0;
    var dict = Object.create(null);
    var chunks = 0;

    while ((pos = aStr.search(DURATION_LETTERS)) !== -1) {
      var type = aStr[pos];
      var numeric = aStr.substr(0, pos);
      aStr = aStr.substr(pos + 1);

      chunks += parseDurationChunk(type, numeric, dict);
    }

    if (chunks < 2) {
      // There must be at least a chunk with "P" and some unit chunk
      throw new Error(
        'invalid duration value: Not enough duration components in "' + aStr + '"'
      );
    }

    return new ICAL.Duration(dict);
  };

  /**
   * Creates a new ICAL.Duration instance from the given data object.
   *
   * @param {Object} aData               An object with members of the duration
   * @param {Number} aData.weeks         Duration in weeks
   * @param {Number} aData.days          Duration in days
   * @param {Number} aData.hours         Duration in hours
   * @param {Number} aData.minutes       Duration in minutes
   * @param {Number} aData.seconds       Duration in seconds
   * @param {Boolean} aData.isNegative   If true, the duration is negative
   * @return {ICAL.Duration}             The createad duration instance
   */
  ICAL.Duration.fromData = function icalduration_from_data(aData) {
    return new ICAL.Duration(aData);
  };
})();
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2012 */



(function() {
  var OPTIONS = ["tzid", "location", "tznames",
                 "latitude", "longitude"];

  /**
   * @classdesc
   * Timezone representation, created by passing in a tzid and component.
   *
   * @example
   * var vcalendar;
   * var timezoneComp = vcalendar.getFirstSubcomponent('vtimezone');
   * var tzid = timezoneComp.getFirstPropertyValue('tzid');
   *
   * var timezone = new ICAL.Timezone({
   *   component: timezoneComp,
   *   tzid
   * });
   *
   * @class
   * @param {ICAL.Component|Object} data options for class
   * @param {String|ICAL.Component} data.component
   *        If data is a simple object, then this member can be set to either a
   *        string containing the component data, or an already parsed
   *        ICAL.Component
   * @param {String} data.tzid      The timezone identifier
   * @param {String} data.location  The timezone locationw
   * @param {String} data.tznames   An alternative string representation of the
   *                                  timezone
   * @param {Number} data.latitude  The latitude of the timezone
   * @param {Number} data.longitude The longitude of the timezone
   */
  ICAL.Timezone = function icaltimezone(data) {
    this.wrappedJSObject = this;
    this.fromData(data);
  };

  ICAL.Timezone.prototype = {

    /**
     * Timezone identifier
     * @type {String}
     */
    tzid: "",

    /**
     * Timezone location
     * @type {String}
     */
    location: "",

    /**
     * Alternative timezone name, for the string representation
     * @type {String}
     */
    tznames: "",

    /**
     * The primary latitude for the timezone.
     * @type {Number}
     */
    latitude: 0.0,

    /**
     * The primary longitude for the timezone.
     * @type {Number}
     */
    longitude: 0.0,

    /**
     * The vtimezone component for this timezone.
     * @type {ICAL.Component}
     */
    component: null,

    /**
     * The year this timezone has been expanded to. All timezone transition
     * dates until this year are known and can be used for calculation
     *
     * @private
     * @type {Number}
     */
    expandedUntilYear: 0,

    /**
     * The class identifier.
     * @constant
     * @type {String}
     * @default "icaltimezone"
     */
    icalclass: "icaltimezone",

    /**
     * Sets up the current instance using members from the passed data object.
     *
     * @param {ICAL.Component|Object} aData options for class
     * @param {String|ICAL.Component} aData.component
     *        If aData is a simple object, then this member can be set to either a
     *        string containing the component data, or an already parsed
     *        ICAL.Component
     * @param {String} aData.tzid      The timezone identifier
     * @param {String} aData.location  The timezone locationw
     * @param {String} aData.tznames   An alternative string representation of the
     *                                  timezone
     * @param {Number} aData.latitude  The latitude of the timezone
     * @param {Number} aData.longitude The longitude of the timezone
     */
    fromData: function fromData(aData) {
      this.expandedUntilYear = 0;
      this.changes = [];

      if (aData instanceof ICAL.Component) {
        // Either a component is passed directly
        this.component = aData;
      } else {
        // Otherwise the component may be in the data object
        if (aData && "component" in aData) {
          if (typeof aData.component == "string") {
            // If a string was passed, parse it as a component
            var jCal = ICAL.parse(aData.component);
            this.component = new ICAL.Component(jCal);
          } else if (aData.component instanceof ICAL.Component) {
            // If it was a component already, then just set it
            this.component = aData.component;
          } else {
            // Otherwise just null out the component
            this.component = null;
          }
        }

        // Copy remaining passed properties
        for (var key in OPTIONS) {
          /* istanbul ignore else */
          if (OPTIONS.hasOwnProperty(key)) {
            var prop = OPTIONS[key];
            if (aData && prop in aData) {
              this[prop] = aData[prop];
            }
          }
        }
      }

      // If we have a component but no TZID, attempt to get it from the
      // component's properties.
      if (this.component instanceof ICAL.Component && !this.tzid) {
        this.tzid = this.component.getFirstPropertyValue('tzid');
      }

      return this;
    },

    /**
     * Finds the utcOffset the given time would occur in this timezone.
     *
     * @param {ICAL.Time} tt        The time to check for
     * @return {Number} utc offset in seconds
     */
    utcOffset: function utcOffset(tt) {
      if (this == ICAL.Timezone.utcTimezone || this == ICAL.Timezone.localTimezone) {
        return 0;
      }

      this._ensureCoverage(tt.year);

      if (!this.changes.length) {
        return 0;
      }

      var tt_change = {
        year: tt.year,
        month: tt.month,
        day: tt.day,
        hour: tt.hour,
        minute: tt.minute,
        second: tt.second
      };

      var change_num = this._findNearbyChange(tt_change);
      var change_num_to_use = -1;
      var step = 1;

      // TODO: replace with bin search?
      for (;;) {
        var change = ICAL.helpers.clone(this.changes[change_num], true);
        if (change.utcOffset < change.prevUtcOffset) {
          ICAL.Timezone.adjust_change(change, 0, 0, 0, change.utcOffset);
        } else {
          ICAL.Timezone.adjust_change(change, 0, 0, 0,
                                          change.prevUtcOffset);
        }

        var cmp = ICAL.Timezone._compare_change_fn(tt_change, change);

        if (cmp >= 0) {
          change_num_to_use = change_num;
        } else {
          step = -1;
        }

        if (step == -1 && change_num_to_use != -1) {
          break;
        }

        change_num += step;

        if (change_num < 0) {
          return 0;
        }

        if (change_num >= this.changes.length) {
          break;
        }
      }

      var zone_change = this.changes[change_num_to_use];
      var utcOffset_change = zone_change.utcOffset - zone_change.prevUtcOffset;

      if (utcOffset_change < 0 && change_num_to_use > 0) {
        var tmp_change = ICAL.helpers.clone(zone_change, true);
        ICAL.Timezone.adjust_change(tmp_change, 0, 0, 0,
                                        tmp_change.prevUtcOffset);

        if (ICAL.Timezone._compare_change_fn(tt_change, tmp_change) < 0) {
          var prev_zone_change = this.changes[change_num_to_use - 1];

          var want_daylight = false; // TODO

          if (zone_change.is_daylight != want_daylight &&
              prev_zone_change.is_daylight == want_daylight) {
            zone_change = prev_zone_change;
          }
        }
      }

      // TODO return is_daylight?
      return zone_change.utcOffset;
    },

    _findNearbyChange: function icaltimezone_find_nearby_change(change) {
      // find the closest match
      var idx = ICAL.helpers.binsearchInsert(
        this.changes,
        change,
        ICAL.Timezone._compare_change_fn
      );

      if (idx >= this.changes.length) {
        return this.changes.length - 1;
      }

      return idx;
    },

    _ensureCoverage: function(aYear) {
      if (ICAL.Timezone._minimumExpansionYear == -1) {
        var today = ICAL.Time.now();
        ICAL.Timezone._minimumExpansionYear = today.year;
      }

      var changesEndYear = aYear;
      if (changesEndYear < ICAL.Timezone._minimumExpansionYear) {
        changesEndYear = ICAL.Timezone._minimumExpansionYear;
      }

      changesEndYear += ICAL.Timezone.EXTRA_COVERAGE;

      if (changesEndYear > ICAL.Timezone.MAX_YEAR) {
        changesEndYear = ICAL.Timezone.MAX_YEAR;
      }

      if (!this.changes.length || this.expandedUntilYear < aYear) {
        var subcomps = this.component.getAllSubcomponents();
        var compLen = subcomps.length;
        var compIdx = 0;

        for (; compIdx < compLen; compIdx++) {
          this._expandComponent(
            subcomps[compIdx], changesEndYear, this.changes
          );
        }

        this.changes.sort(ICAL.Timezone._compare_change_fn);
        this.expandedUntilYear = changesEndYear;
      }
    },

    _expandComponent: function(aComponent, aYear, changes) {
      if (!aComponent.hasProperty("dtstart") ||
          !aComponent.hasProperty("tzoffsetto") ||
          !aComponent.hasProperty("tzoffsetfrom")) {
        return null;
      }

      var dtstart = aComponent.getFirstProperty("dtstart").getFirstValue();
      var change;

      function convert_tzoffset(offset) {
        return offset.factor * (offset.hours * 3600 + offset.minutes * 60);
      }

      function init_changes() {
        var changebase = {};
        changebase.is_daylight = (aComponent.name == "daylight");
        changebase.utcOffset = convert_tzoffset(
          aComponent.getFirstProperty("tzoffsetto").getFirstValue()
        );

        changebase.prevUtcOffset = convert_tzoffset(
          aComponent.getFirstProperty("tzoffsetfrom").getFirstValue()
        );

        return changebase;
      }

      if (!aComponent.hasProperty("rrule") && !aComponent.hasProperty("rdate")) {
        change = init_changes();
        change.year = dtstart.year;
        change.month = dtstart.month;
        change.day = dtstart.day;
        change.hour = dtstart.hour;
        change.minute = dtstart.minute;
        change.second = dtstart.second;

        ICAL.Timezone.adjust_change(change, 0, 0, 0,
                                        -change.prevUtcOffset);
        changes.push(change);
      } else {
        var props = aComponent.getAllProperties("rdate");
        for (var rdatekey in props) {
          /* istanbul ignore if */
          if (!props.hasOwnProperty(rdatekey)) {
            continue;
          }
          var rdate = props[rdatekey];
          var time = rdate.getFirstValue();
          change = init_changes();

          change.year = time.year;
          change.month = time.month;
          change.day = time.day;

          if (time.isDate) {
            change.hour = dtstart.hour;
            change.minute = dtstart.minute;
            change.second = dtstart.second;

            if (dtstart.zone != ICAL.Timezone.utcTimezone) {
              ICAL.Timezone.adjust_change(change, 0, 0, 0,
                                              -change.prevUtcOffset);
            }
          } else {
            change.hour = time.hour;
            change.minute = time.minute;
            change.second = time.second;

            if (time.zone != ICAL.Timezone.utcTimezone) {
              ICAL.Timezone.adjust_change(change, 0, 0, 0,
                                              -change.prevUtcOffset);
            }
          }

          changes.push(change);
        }

        var rrule = aComponent.getFirstProperty("rrule");

        if (rrule) {
          rrule = rrule.getFirstValue();
          change = init_changes();

          if (rrule.until && rrule.until.zone == ICAL.Timezone.utcTimezone) {
            rrule.until.adjust(0, 0, 0, change.prevUtcOffset);
            rrule.until.zone = ICAL.Timezone.localTimezone;
          }

          var iterator = rrule.iterator(dtstart);

          var occ;
          while ((occ = iterator.next())) {
            change = init_changes();
            if (occ.year > aYear || !occ) {
              break;
            }

            change.year = occ.year;
            change.month = occ.month;
            change.day = occ.day;
            change.hour = occ.hour;
            change.minute = occ.minute;
            change.second = occ.second;
            change.isDate = occ.isDate;

            ICAL.Timezone.adjust_change(change, 0, 0, 0,
                                            -change.prevUtcOffset);
            changes.push(change);
          }
        }
      }

      return changes;
    },

    /**
     * The string representation of this timezone.
     * @return {String}
     */
    toString: function toString() {
      return (this.tznames ? this.tznames : this.tzid);
    }
  };

  ICAL.Timezone._compare_change_fn = function icaltimezone_compare_change_fn(a, b) {
    if (a.year < b.year) return -1;
    else if (a.year > b.year) return 1;

    if (a.month < b.month) return -1;
    else if (a.month > b.month) return 1;

    if (a.day < b.day) return -1;
    else if (a.day > b.day) return 1;

    if (a.hour < b.hour) return -1;
    else if (a.hour > b.hour) return 1;

    if (a.minute < b.minute) return -1;
    else if (a.minute > b.minute) return 1;

    if (a.second < b.second) return -1;
    else if (a.second > b.second) return 1;

    return 0;
  };

  /**
   * Convert the date/time from one zone to the next.
   *
   * @param {ICAL.Time} tt                  The time to convert
   * @param {ICAL.Timezone} from_zone       The source zone to convert from
   * @param {ICAL.Timezone} to_zone         The target zone to convert to
   * @return {ICAL.Time}                    The converted date/time object
   */
  ICAL.Timezone.convert_time = function icaltimezone_convert_time(tt, from_zone, to_zone) {
    if (tt.isDate ||
        from_zone.tzid == to_zone.tzid ||
        from_zone == ICAL.Timezone.localTimezone ||
        to_zone == ICAL.Timezone.localTimezone) {
      tt.zone = to_zone;
      return tt;
    }

    var utcOffset = from_zone.utcOffset(tt);
    tt.adjust(0, 0, 0, - utcOffset);

    utcOffset = to_zone.utcOffset(tt);
    tt.adjust(0, 0, 0, utcOffset);

    return null;
  };

  /**
   * Creates a new ICAL.Timezone instance from the passed data object.
   *
   * @param {ICAL.Component|Object} aData options for class
   * @param {String|ICAL.Component} aData.component
   *        If aData is a simple object, then this member can be set to either a
   *        string containing the component data, or an already parsed
   *        ICAL.Component
   * @param {String} aData.tzid      The timezone identifier
   * @param {String} aData.location  The timezone locationw
   * @param {String} aData.tznames   An alternative string representation of the
   *                                  timezone
   * @param {Number} aData.latitude  The latitude of the timezone
   * @param {Number} aData.longitude The longitude of the timezone
   */
  ICAL.Timezone.fromData = function icaltimezone_fromData(aData) {
    var tt = new ICAL.Timezone();
    return tt.fromData(aData);
  };

  /**
   * The instance describing the UTC timezone
   * @type {ICAL.Timezone}
   * @constant
   * @instance
   */
  ICAL.Timezone.utcTimezone = ICAL.Timezone.fromData({
    tzid: "UTC"
  });

  /**
   * The instance describing the local timezone
   * @type {ICAL.Timezone}
   * @constant
   * @instance
   */
  ICAL.Timezone.localTimezone = ICAL.Timezone.fromData({
    tzid: "floating"
  });

  /**
   * Adjust a timezone change object.
   * @private
   * @param {Object} change     The timezone change object
   * @param {Number} days       The extra amount of days
   * @param {Number} hours      The extra amount of hours
   * @param {Number} minutes    The extra amount of minutes
   * @param {Number} seconds    The extra amount of seconds
   */
  ICAL.Timezone.adjust_change = function icaltimezone_adjust_change(change, days, hours, minutes, seconds) {
    return ICAL.Time.prototype.adjust.call(
      change,
      days,
      hours,
      minutes,
      seconds,
      change
    );
  };

  ICAL.Timezone._minimumExpansionYear = -1;
  ICAL.Timezone.MAX_YEAR = 2035; // TODO this is because of time_t, which we don't need. Still usefull?
  ICAL.Timezone.EXTRA_COVERAGE = 5;
})();
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.TimezoneService = (function() {
  var zones;

  /**
   * @classdesc
   * Singleton class to contain timezones.  Right now its all manual registry in
   * the future we may use this class to download timezone information or handle
   * loading pre-expanded timezones.
   *
   * @namespace
   * @alias ICAL.TimezoneService
   */
  var TimezoneService = {
    get count() {
      return Object.keys(zones).length;
    },

    reset: function() {
      zones = Object.create(null);
      var utc = ICAL.Timezone.utcTimezone;

      zones.Z = utc;
      zones.UTC = utc;
      zones.GMT = utc;
    },

    /**
     * Checks if timezone id has been registered.
     *
     * @param {String} tzid     Timezone identifier (e.g. America/Los_Angeles)
     * @return {Boolean}        False, when not present
     */
    has: function(tzid) {
      return !!zones[tzid];
    },

    /**
     * Returns a timezone by its tzid if present.
     *
     * @param {String} tzid     Timezone identifier (e.g. America/Los_Angeles)
     * @return {?ICAL.Timezone} The timezone, or null if not found
     */
    get: function(tzid) {
      return zones[tzid];
    },

    /**
     * Registers a timezone object or component.
     *
     * @param {String=} name
     *        The name of the timezone. Defaults to the component's TZID if not
     *        passed.
     * @param {ICAL.Component|ICAL.Timezone} zone
     *        The initialized zone or vtimezone.
     */
    register: function(name, timezone) {
      if (name instanceof ICAL.Component) {
        if (name.name === 'vtimezone') {
          timezone = new ICAL.Timezone(name);
          name = timezone.tzid;
        }
      }

      if (timezone instanceof ICAL.Timezone) {
        zones[name] = timezone;
      } else {
        throw new TypeError('timezone must be ICAL.Timezone or ICAL.Component');
      }
    },

    /**
     * Removes a timezone by its tzid from the list.
     *
     * @param {String} tzid     Timezone identifier (e.g. America/Los_Angeles)
     * @return {?ICAL.Timezone} The removed timezone, or null if not registered
     */
    remove: function(tzid) {
      return (delete zones[tzid]);
    }
  };

  // initialize defaults
  TimezoneService.reset();

  return TimezoneService;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */



(function() {

  /**
   * @classdesc
   * iCalendar Time representation (similar to JS Date object).  Fully
   * independent of system (OS) timezone / time.  Unlike JS Date, the month
   * January is 1, not zero.
   *
   * @example
   * var time = new ICAL.Time({
   *   year: 2012,
   *   month: 10,
   *   day: 11
   *   minute: 0,
   *   second: 0,
   *   isDate: false
   * });
   *
   *
   * @alias ICAL.Time
   * @class
   * @param {Object} data           Time initialization
   * @param {Number=} data.year     The year for this date
   * @param {Number=} data.month    The month for this date
   * @param {Number=} data.day      The day for this date
   * @param {Number=} data.hour     The hour for this date
   * @param {Number=} data.minute   The minute for this date
   * @param {Number=} data.second   The second for this date
   * @param {Boolean=} data.isDate  If true, the instance represents a date (as
   *                                  opposed to a date-time)
   * @param {ICAL.Timezone} zone timezone this position occurs in
   */
  ICAL.Time = function icaltime(data, zone) {
    this.wrappedJSObject = this;
    var time = this._time = Object.create(null);

    /* time defaults */
    time.year = 0;
    time.month = 1;
    time.day = 1;
    time.hour = 0;
    time.minute = 0;
    time.second = 0;
    time.isDate = false;

    this.fromData(data, zone);
  };

  ICAL.Time._dowCache = {};
  ICAL.Time._wnCache = {};

  ICAL.Time.prototype = {

    /**
     * The class identifier.
     * @constant
     * @type {String}
     * @default "icaltime"
     */
    icalclass: "icaltime",
    _cachedUnixTime: null,

    /**
     * The type name, to be used in the jCal object. This value may change and
     * is strictly defined by the {@link ICAL.Time#isDate isDate} member.
     * @readonly
     * @type {String}
     * @default "date-time"
     */
    get icaltype() {
      return this.isDate ? 'date' : 'date-time';
    },

    /**
     * The timezone for this time.
     * @type {ICAL.Timezone}
     */
    zone: null,

    /**
     * Internal uses to indicate that a change has been made and the next read
     * operation must attempt to normalize the value (for example changing the
     * day to 33).
     *
     * @type {Boolean}
     * @private
     */
    _pendingNormalization: false,

    /**
     * Returns a clone of the time object.
     *
     * @return {ICAL.Time}              The cloned object
     */
    clone: function() {
      return new ICAL.Time(this._time, this.zone);
    },

    /**
     * Reset the time instance to epoch time
     */
    reset: function icaltime_reset() {
      this.fromData(ICAL.Time.epochTime);
      this.zone = ICAL.Timezone.utcTimezone;
    },

    /**
     * Reset the time instance to the given date/time values.
     *
     * @param {Number} year             The year to set
     * @param {Number} month            The month to set
     * @param {Number} day              The day to set
     * @param {Number} hour             The hour to set
     * @param {Number} minute           The minute to set
     * @param {Number} second           The second to set
     * @param {ICAL.Timezone} timezone  The timezone to set
     */
    resetTo: function icaltime_resetTo(year, month, day,
                                       hour, minute, second, timezone) {
      this.fromData({
        year: year,
        month: month,
        day: day,
        hour: hour,
        minute: minute,
        second: second,
        zone: timezone
      });
    },

    /**
     * Set up the current instance from the Javascript date value.
     *
     * @param {?Date} aDate     The Javascript Date to read, or null to reset
     * @param {Boolean} useUTC  If true, the UTC values of the date will be used
     */
    fromJSDate: function icaltime_fromJSDate(aDate, useUTC) {
      if (!aDate) {
        this.reset();
      } else {
        if (useUTC) {
          this.zone = ICAL.Timezone.utcTimezone;
          this.year = aDate.getUTCFullYear();
          this.month = aDate.getUTCMonth() + 1;
          this.day = aDate.getUTCDate();
          this.hour = aDate.getUTCHours();
          this.minute = aDate.getUTCMinutes();
          this.second = aDate.getUTCSeconds();
        } else {
          this.zone = ICAL.Timezone.localTimezone;
          this.year = aDate.getFullYear();
          this.month = aDate.getMonth() + 1;
          this.day = aDate.getDate();
          this.hour = aDate.getHours();
          this.minute = aDate.getMinutes();
          this.second = aDate.getSeconds();
        }
      }
      this._cachedUnixTime = null;
      return this;
    },

    /**
     * Sets up the current instance using members from the passed data object.
     *
     * @param {Object} aData            Time initialization
     * @param {Number=} aData.year      The year for this date
     * @param {Number=} aData.month     The month for this date
     * @param {Number=} aData.day       The day for this date
     * @param {Number=} aData.hour      The hour for this date
     * @param {Number=} aData.minute    The minute for this date
     * @param {Number=} aData.second    The second for this date
     * @param {Boolean=} aData.isDate   If true, the instance represents a date
     *                                    (as opposed to a date-time)
     * @param {ICAL.Timezone=} aZone    Timezone this position occurs in
     */
    fromData: function fromData(aData, aZone) {
      if (aData) {
        for (var key in aData) {
          /* istanbul ignore else */
          if (Object.prototype.hasOwnProperty.call(aData, key)) {
            // ical type cannot be set
            if (key === 'icaltype') continue;
            this[key] = aData[key];
          }
        }
      }

      if (aZone) {
        this.zone = aZone;
      }

      if (aData && !("isDate" in aData)) {
        this.isDate = !("hour" in aData);
      } else if (aData && ("isDate" in aData)) {
        this.isDate = aData.isDate;
      }

      if (aData && "timezone" in aData) {
        var zone = ICAL.TimezoneService.get(
          aData.timezone
        );

        this.zone = zone || ICAL.Timezone.localTimezone;
      }

      if (aData && "zone" in aData) {
        this.zone = aData.zone;
      }

      if (!this.zone) {
        this.zone = ICAL.Timezone.localTimezone;
      }

      this._cachedUnixTime = null;
      return this;
    },

    /**
     * Calculate the day of week.
     * @param {ICAL.Time.weekDay=} aWeekStart
     *        The week start weekday, defaults to SUNDAY
     * @return {ICAL.Time.weekDay}
     */
    dayOfWeek: function icaltime_dayOfWeek(aWeekStart) {
      var firstDow = aWeekStart || ICAL.Time.SUNDAY;
      var dowCacheKey = (this.year << 12) + (this.month << 8) + (this.day << 3) + firstDow;
      if (dowCacheKey in ICAL.Time._dowCache) {
        return ICAL.Time._dowCache[dowCacheKey];
      }

      // Using Zeller's algorithm
      var q = this.day;
      var m = this.month + (this.month < 3 ? 12 : 0);
      var Y = this.year - (this.month < 3 ? 1 : 0);

      var h = (q + Y + ICAL.helpers.trunc(((m + 1) * 26) / 10) + ICAL.helpers.trunc(Y / 4));
      /* istanbul ignore else */
      if (true /* gregorian */) {
        h += ICAL.helpers.trunc(Y / 100) * 6 + ICAL.helpers.trunc(Y / 400);
      } else {
        h += 5;
      }

      // Normalize to 1 = wkst
      h = ((h + 7 - firstDow) % 7) + 1;
      ICAL.Time._dowCache[dowCacheKey] = h;
      return h;
    },

    /**
     * Calculate the day of year.
     * @return {Number}
     */
    dayOfYear: function dayOfYear() {
      var is_leap = (ICAL.Time.isLeapYear(this.year) ? 1 : 0);
      var diypm = ICAL.Time.daysInYearPassedMonth;
      return diypm[is_leap][this.month - 1] + this.day;
    },

    /**
     * Returns a copy of the current date/time, rewound to the start of the
     * week. The resulting ICAL.Time instance is of icaltype date, even if this
     * is a date-time.
     *
     * @param {ICAL.Time.weekDay=} aWeekStart
     *        The week start weekday, defaults to SUNDAY
     * @return {ICAL.Time}      The start of the week (cloned)
     */
    startOfWeek: function startOfWeek(aWeekStart) {
      var firstDow = aWeekStart || ICAL.Time.SUNDAY;
      var result = this.clone();
      result.day -= ((this.dayOfWeek() + 7 - firstDow) % 7);
      result.isDate = true;
      result.hour = 0;
      result.minute = 0;
      result.second = 0;
      return result;
    },

    /**
     * Returns a copy of the current date/time, shifted to the end of the week.
     * The resulting ICAL.Time instance is of icaltype date, even if this is a
     * date-time.
     *
     * @param {ICAL.Time.weekDay=} aWeekStart
     *        The week start weekday, defaults to SUNDAY
     * @return {ICAL.Time}      The end of the week (cloned)
     */
    endOfWeek: function endOfWeek(aWeekStart) {
      var firstDow = aWeekStart || ICAL.Time.SUNDAY;
      var result = this.clone();
      result.day += (7 - this.dayOfWeek() + firstDow - ICAL.Time.SUNDAY) % 7;
      result.isDate = true;
      result.hour = 0;
      result.minute = 0;
      result.second = 0;
      return result;
    },

    /**
     * Returns a copy of the current date/time, rewound to the start of the
     * month. The resulting ICAL.Time instance is of icaltype date, even if
     * this is a date-time.
     *
     * @return {ICAL.Time}      The start of the month (cloned)
     */
    startOfMonth: function startOfMonth() {
      var result = this.clone();
      result.day = 1;
      result.isDate = true;
      result.hour = 0;
      result.minute = 0;
      result.second = 0;
      return result;
    },

    /**
     * Returns a copy of the current date/time, shifted to the end of the
     * month.  The resulting ICAL.Time instance is of icaltype date, even if
     * this is a date-time.
     *
     * @return {ICAL.Time}      The end of the month (cloned)
     */
    endOfMonth: function endOfMonth() {
      var result = this.clone();
      result.day = ICAL.Time.daysInMonth(result.month, result.year);
      result.isDate = true;
      result.hour = 0;
      result.minute = 0;
      result.second = 0;
      return result;
    },

    /**
     * Returns a copy of the current date/time, rewound to the start of the
     * year. The resulting ICAL.Time instance is of icaltype date, even if
     * this is a date-time.
     *
     * @return {ICAL.Time}      The start of the year (cloned)
     */
    startOfYear: function startOfYear() {
      var result = this.clone();
      result.day = 1;
      result.month = 1;
      result.isDate = true;
      result.hour = 0;
      result.minute = 0;
      result.second = 0;
      return result;
    },

    /**
     * Returns a copy of the current date/time, shifted to the end of the
     * year.  The resulting ICAL.Time instance is of icaltype date, even if
     * this is a date-time.
     *
     * @return {ICAL.Time}      The end of the year (cloned)
     */
    endOfYear: function endOfYear() {
      var result = this.clone();
      result.day = 31;
      result.month = 12;
      result.isDate = true;
      result.hour = 0;
      result.minute = 0;
      result.second = 0;
      return result;
    },

    /**
     * First calculates the start of the week, then returns the day of year for
     * this date. If the day falls into the previous year, the day is zero or negative.
     *
     * @param {ICAL.Time.weekDay=} aFirstDayOfWeek
     *        The week start weekday, defaults to SUNDAY
     * @return {Number}     The calculated day of year
     */
    startDoyWeek: function startDoyWeek(aFirstDayOfWeek) {
      var firstDow = aFirstDayOfWeek || ICAL.Time.SUNDAY;
      var delta = this.dayOfWeek() - firstDow;
      if (delta < 0) delta += 7;
      return this.dayOfYear() - delta;
    },

    /**
     * Get the dominical letter for the current year. Letters range from A - G
     * for common years, and AG to GF for leap years.
     *
     * @param {Number} yr           The year to retrieve the letter for
     * @return {String}             The dominical letter.
     */
    getDominicalLetter: function() {
      return ICAL.Time.getDominicalLetter(this.year);
    },

    /**
     * Finds the nthWeekDay relative to the current month (not day).  The
     * returned value is a day relative the month that this month belongs to so
     * 1 would indicate the first of the month and 40 would indicate a day in
     * the following month.
     *
     * @param {Number} aDayOfWeek   Day of the week see the day name constants
     * @param {Number} aPos         Nth occurrence of a given week day values
     *        of 1 and 0 both indicate the first weekday of that type. aPos may
     *        be either positive or negative
     *
     * @return {Number} numeric value indicating a day relative
     *                   to the current month of this time object
     */
    nthWeekDay: function icaltime_nthWeekDay(aDayOfWeek, aPos) {
      var daysInMonth = ICAL.Time.daysInMonth(this.month, this.year);
      var weekday;
      var pos = aPos;

      var start = 0;

      var otherDay = this.clone();

      if (pos >= 0) {
        otherDay.day = 1;

        // because 0 means no position has been given
        // 1 and 0 indicate the same day.
        if (pos != 0) {
          // remove the extra numeric value
          pos--;
        }

        // set current start offset to current day.
        start = otherDay.day;

        // find the current day of week
        var startDow = otherDay.dayOfWeek();

        // calculate the difference between current
        // day of the week and desired day of the week
        var offset = aDayOfWeek - startDow;


        // if the offset goes into the past
        // week we add 7 so its goes into the next
        // week. We only want to go forward in time here.
        if (offset < 0)
          // this is really important otherwise we would
          // end up with dates from in the past.
          offset += 7;

        // add offset to start so start is the same
        // day of the week as the desired day of week.
        start += offset;

        // because we are going to add (and multiply)
        // the numeric value of the day we subtract it
        // from the start position so not to add it twice.
        start -= aDayOfWeek;

        // set week day
        weekday = aDayOfWeek;
      } else {

        // then we set it to the last day in the current month
        otherDay.day = daysInMonth;

        // find the ends weekday
        var endDow = otherDay.dayOfWeek();

        pos++;

        weekday = (endDow - aDayOfWeek);

        if (weekday < 0) {
          weekday += 7;
        }

        weekday = daysInMonth - weekday;
      }

      weekday += pos * 7;

      return start + weekday;
    },

    /**
     * Checks if current time is the nth weekday, relative to the current
     * month.  Will always return false when rule resolves outside of current
     * month.
     *
     * @param {ICAL.Time.weekDay} aDayOfWeek       Day of week to check
     * @param {Number} aPos                        Relative position
     * @return {Boolean}                           True, if its the nth weekday
     */
    isNthWeekDay: function(aDayOfWeek, aPos) {
      var dow = this.dayOfWeek();

      if (aPos === 0 && dow === aDayOfWeek) {
        return true;
      }

      // get pos
      var day = this.nthWeekDay(aDayOfWeek, aPos);

      if (day === this.day) {
        return true;
      }

      return false;
    },

    /**
     * Calculates the ISO 8601 week number. The first week of a year is the
     * week that contains the first Thursday. The year can have 53 weeks, if
     * January 1st is a Friday.
     *
     * Note there are regions where the first week of the year is the one that
     * starts on January 1st, which may offset the week number. Also, if a
     * different week start is specified, this will also affect the week
     * number.
     *
     * @see ICAL.Time.weekOneStarts
     * @param {ICAL.Time.weekDay} aWeekStart        The weekday the week starts with
     * @return {Number}                             The ISO week number
     */
    weekNumber: function weekNumber(aWeekStart) {
      var wnCacheKey = (this.year << 12) + (this.month << 8) + (this.day << 3) + aWeekStart;
      if (wnCacheKey in ICAL.Time._wnCache) {
        return ICAL.Time._wnCache[wnCacheKey];
      }
      // This function courtesty of Julian Bucknall, published under the MIT license
      // http://www.boyet.com/articles/publishedarticles/calculatingtheisoweeknumb.html
      // plus some fixes to be able to use different week starts.
      var week1;

      var dt = this.clone();
      dt.isDate = true;
      var isoyear = this.year;

      if (dt.month == 12 && dt.day > 25) {
        week1 = ICAL.Time.weekOneStarts(isoyear + 1, aWeekStart);
        if (dt.compare(week1) < 0) {
          week1 = ICAL.Time.weekOneStarts(isoyear, aWeekStart);
        } else {
          isoyear++;
        }
      } else {
        week1 = ICAL.Time.weekOneStarts(isoyear, aWeekStart);
        if (dt.compare(week1) < 0) {
          week1 = ICAL.Time.weekOneStarts(--isoyear, aWeekStart);
        }
      }

      var daysBetween = (dt.subtractDate(week1).toSeconds() / 86400);
      var answer = ICAL.helpers.trunc(daysBetween / 7) + 1;
      ICAL.Time._wnCache[wnCacheKey] = answer;
      return answer;
    },

    /**
     * Adds the duration to the current time. The instance is modified in
     * place.
     *
     * @param {ICAL.Duration} aDuration         The duration to add
     */
    addDuration: function icaltime_add(aDuration) {
      var mult = (aDuration.isNegative ? -1 : 1);

      // because of the duration optimizations it is much
      // more efficient to grab all the values up front
      // then set them directly (which will avoid a normalization call).
      // So we don't actually normalize until we need it.
      var second = this.second;
      var minute = this.minute;
      var hour = this.hour;
      var day = this.day;

      second += mult * aDuration.seconds;
      minute += mult * aDuration.minutes;
      hour += mult * aDuration.hours;
      day += mult * aDuration.days;
      day += mult * 7 * aDuration.weeks;

      this.second = second;
      this.minute = minute;
      this.hour = hour;
      this.day = day;

      this._cachedUnixTime = null;
    },

    /**
     * Subtract the date details (_excluding_ timezone).  Useful for finding
     * the relative difference between two time objects excluding their
     * timezone differences.
     *
     * @param {ICAL.Time} aDate     The date to substract
     * @return {ICAL.Duration}      The difference as a duration
     */
    subtractDate: function icaltime_subtract(aDate) {
      var unixTime = this.toUnixTime() + this.utcOffset();
      var other = aDate.toUnixTime() + aDate.utcOffset();
      return ICAL.Duration.fromSeconds(unixTime - other);
    },

    /**
     * Subtract the date details, taking timezones into account.
     *
     * @param {ICAL.Time} aDate  The date to subtract
     * @return {ICAL.Duration}  The difference in duration
     */
    subtractDateTz: function icaltime_subtract_abs(aDate) {
      var unixTime = this.toUnixTime();
      var other = aDate.toUnixTime();
      return ICAL.Duration.fromSeconds(unixTime - other);
    },

    /**
     * Compares the ICAL.Time instance with another one.
     *
     * @param {ICAL.Duration} aOther        The instance to compare with
     * @return {Number}                     -1, 0 or 1 for less/equal/greater
     */
    compare: function icaltime_compare(other) {
      var a = this.toUnixTime();
      var b = other.toUnixTime();

      if (a > b) return 1;
      if (b > a) return -1;
      return 0;
    },

    /**
     * Compares only the date part of this instance with another one.
     *
     * @param {ICAL.Duration} other         The instance to compare with
     * @param {ICAL.Timezone} tz            The timezone to compare in
     * @return {Number}                     -1, 0 or 1 for less/equal/greater
     */
    compareDateOnlyTz: function icaltime_compareDateOnlyTz(other, tz) {
      function cmp(attr) {
        return ICAL.Time._cmp_attr(a, b, attr);
      }
      var a = this.convertToZone(tz);
      var b = other.convertToZone(tz);
      var rc = 0;

      if ((rc = cmp("year")) != 0) return rc;
      if ((rc = cmp("month")) != 0) return rc;
      if ((rc = cmp("day")) != 0) return rc;

      return rc;
    },

    /**
     * Convert the instance into another timzone. The returned ICAL.Time
     * instance is always a copy.
     *
     * @param {ICAL.Timezone} zone      The zone to convert to
     * @return {ICAL.Time}              The copy, converted to the zone
     */
    convertToZone: function convertToZone(zone) {
      var copy = this.clone();
      var zone_equals = (this.zone.tzid == zone.tzid);

      if (!this.isDate && !zone_equals) {
        ICAL.Timezone.convert_time(copy, this.zone, zone);
      }

      copy.zone = zone;
      return copy;
    },

    /**
     * Calculates the UTC offset of the current date/time in the timezone it is
     * in.
     *
     * @return {Number}     UTC offset in seconds
     */
    utcOffset: function utc_offset() {
      if (this.zone == ICAL.Timezone.localTimezone ||
          this.zone == ICAL.Timezone.utcTimezone) {
        return 0;
      } else {
        return this.zone.utcOffset(this);
      }
    },

    /**
     * Returns an RFC 5545 compliant ical representation of this object.
     *
     * @return {String} ical date/date-time
     */
    toICALString: function() {
      var string = this.toString();

      if (string.length > 10) {
        return ICAL.design.icalendar.value['date-time'].toICAL(string);
      } else {
        return ICAL.design.icalendar.value.date.toICAL(string);
      }
    },

    /**
     * The string representation of this date/time, in jCal form
     * (including : and - separators).
     * @return {String}
     */
    toString: function toString() {
      var result = this.year + '-' +
                   ICAL.helpers.pad2(this.month) + '-' +
                   ICAL.helpers.pad2(this.day);

      if (!this.isDate) {
          result += 'T' + ICAL.helpers.pad2(this.hour) + ':' +
                    ICAL.helpers.pad2(this.minute) + ':' +
                    ICAL.helpers.pad2(this.second);

        if (this.zone === ICAL.Timezone.utcTimezone) {
          result += 'Z';
        }
      }

      return result;
    },

    /**
     * Converts the current instance to a Javascript date
     * @return {Date}
     */
    toJSDate: function toJSDate() {
      if (this.zone == ICAL.Timezone.localTimezone) {
        if (this.isDate) {
          return new Date(this.year, this.month - 1, this.day);
        } else {
          return new Date(this.year, this.month - 1, this.day,
                          this.hour, this.minute, this.second, 0);
        }
      } else {
        return new Date(this.toUnixTime() * 1000);
      }
    },

    _normalize: function icaltime_normalize() {
      var isDate = this._time.isDate;
      if (this._time.isDate) {
        this._time.hour = 0;
        this._time.minute = 0;
        this._time.second = 0;
      }
      this.adjust(0, 0, 0, 0);

      return this;
    },

    /**
     * Adjust the date/time by the given offset
     *
     * @param {Number} aExtraDays       The extra amount of days
     * @param {Number} aExtraHours      The extra amount of hours
     * @param {Number} aExtraMinutes    The extra amount of minutes
     * @param {Number} aExtraSeconds    The extra amount of seconds
     * @param {Number=} aTime           The time to adjust, defaults to the
     *                                    current instance.
     */
    adjust: function icaltime_adjust(aExtraDays, aExtraHours,
                                     aExtraMinutes, aExtraSeconds, aTime) {

      var minutesOverflow, hoursOverflow,
          daysOverflow = 0, yearsOverflow = 0;

      var second, minute, hour, day;
      var daysInMonth;

      var time = aTime || this._time;

      if (!time.isDate) {
        second = time.second + aExtraSeconds;
        time.second = second % 60;
        minutesOverflow = ICAL.helpers.trunc(second / 60);
        if (time.second < 0) {
          time.second += 60;
          minutesOverflow--;
        }

        minute = time.minute + aExtraMinutes + minutesOverflow;
        time.minute = minute % 60;
        hoursOverflow = ICAL.helpers.trunc(minute / 60);
        if (time.minute < 0) {
          time.minute += 60;
          hoursOverflow--;
        }

        hour = time.hour + aExtraHours + hoursOverflow;

        time.hour = hour % 24;
        daysOverflow = ICAL.helpers.trunc(hour / 24);
        if (time.hour < 0) {
          time.hour += 24;
          daysOverflow--;
        }
      }


      // Adjust month and year first, because we need to know what month the day
      // is in before adjusting it.
      if (time.month > 12) {
        yearsOverflow = ICAL.helpers.trunc((time.month - 1) / 12);
      } else if (time.month < 1) {
        yearsOverflow = ICAL.helpers.trunc(time.month / 12) - 1;
      }

      time.year += yearsOverflow;
      time.month -= 12 * yearsOverflow;

      // Now take care of the days (and adjust month if needed)
      day = time.day + aExtraDays + daysOverflow;

      if (day > 0) {
        for (;;) {
          daysInMonth = ICAL.Time.daysInMonth(time.month, time.year);
          if (day <= daysInMonth) {
            break;
          }

          time.month++;
          if (time.month > 12) {
            time.year++;
            time.month = 1;
          }

          day -= daysInMonth;
        }
      } else {
        while (day <= 0) {
          if (time.month == 1) {
            time.year--;
            time.month = 12;
          } else {
            time.month--;
          }

          day += ICAL.Time.daysInMonth(time.month, time.year);
        }
      }

      time.day = day;

      this._cachedUnixTime = null;
      return this;
    },

    /**
     * Sets up the current instance from unix time, the number of seconds since
     * January 1st, 1970.
     *
     * @param {Number} seconds      The seconds to set up with
     */
    fromUnixTime: function fromUnixTime(seconds) {
      this.zone = ICAL.Timezone.utcTimezone;
      var epoch = ICAL.Time.epochTime.clone();
      epoch.adjust(0, 0, 0, seconds);

      this.year = epoch.year;
      this.month = epoch.month;
      this.day = epoch.day;
      this.hour = epoch.hour;
      this.minute = epoch.minute;
      this.second = Math.floor(epoch.second);

      this._cachedUnixTime = null;
    },

    /**
     * Converts the current instance to seconds since January 1st 1970.
     *
     * @return {Number}         Seconds since 1970
     */
    toUnixTime: function toUnixTime() {
      if (this._cachedUnixTime !== null) {
        return this._cachedUnixTime;
      }
      var offset = this.utcOffset();

      // we use the offset trick to ensure
      // that we are getting the actual UTC time
      var ms = Date.UTC(
        this.year,
        this.month - 1,
        this.day,
        this.hour,
        this.minute,
        this.second - offset
      );

      // seconds
      this._cachedUnixTime = ms / 1000;
      return this._cachedUnixTime;
    },

    /**
     * Converts time to into Object which can be serialized then re-created
     * using the constructor.
     *
     * @example
     * // toJSON will automatically be called
     * var json = JSON.stringify(mytime);
     *
     * var deserialized = JSON.parse(json);
     *
     * var time = new ICAL.Time(deserialized);
     *
     * @return {Object}
     */
    toJSON: function() {
      var copy = [
        'year',
        'month',
        'day',
        'hour',
        'minute',
        'second',
        'isDate'
      ];

      var result = Object.create(null);

      var i = 0;
      var len = copy.length;
      var prop;

      for (; i < len; i++) {
        prop = copy[i];
        result[prop] = this[prop];
      }

      if (this.zone) {
        result.timezone = this.zone.tzid;
      }

      return result;
    }

  };

  (function setupNormalizeAttributes() {
    // This needs to run before any instances are created!
    function defineAttr(attr) {
      Object.defineProperty(ICAL.Time.prototype, attr, {
        get: function getTimeAttr() {
          if (this._pendingNormalization) {
            this._normalize();
            this._pendingNormalization = false;
          }

          return this._time[attr];
        },
        set: function setTimeAttr(val) {
          // Check if isDate will be set and if was not set to normalize date.
          // This avoids losing days when seconds, minutes and hours are zeroed
          // what normalize will do when time is a date.
          if (attr === "isDate" && val && !this._time.isDate) {
            this.adjust(0, 0, 0, 0);
          }
          this._cachedUnixTime = null;
          this._pendingNormalization = true;
          this._time[attr] = val;

          return val;
        }
      });

    }

    /* istanbul ignore else */
    if ("defineProperty" in Object) {
      defineAttr("year");
      defineAttr("month");
      defineAttr("day");
      defineAttr("hour");
      defineAttr("minute");
      defineAttr("second");
      defineAttr("isDate");
    }
  })();

  /**
   * Returns the days in the given month
   *
   * @param {Number} month      The month to check
   * @param {Number} year       The year to check
   * @return {Number}           The number of days in the month
   */
  ICAL.Time.daysInMonth = function icaltime_daysInMonth(month, year) {
    var _daysInMonth = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    var days = 30;

    if (month < 1 || month > 12) return days;

    days = _daysInMonth[month];

    if (month == 2) {
      days += ICAL.Time.isLeapYear(year);
    }

    return days;
  };

  /**
   * Checks if the year is a leap year
   *
   * @param {Number} year       The year to check
   * @return {Boolean}          True, if the year is a leap year
   */
  ICAL.Time.isLeapYear = function isLeapYear(year) {
    if (year <= 1752) {
      return ((year % 4) == 0);
    } else {
      return (((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0));
    }
  };

  /**
   * Create a new ICAL.Time from the day of year and year. The date is returned
   * in floating timezone.
   *
   * @param {Number} aDayOfYear     The day of year
   * @param {Number} aYear          The year to create the instance in
   * @return {ICAL.Time}            The created instance with the calculated date
   */
  ICAL.Time.fromDayOfYear = function icaltime_fromDayOfYear(aDayOfYear, aYear) {
    var year = aYear;
    var doy = aDayOfYear;
    var tt = new ICAL.Time();
    tt.auto_normalize = false;
    var is_leap = (ICAL.Time.isLeapYear(year) ? 1 : 0);

    if (doy < 1) {
      year--;
      is_leap = (ICAL.Time.isLeapYear(year) ? 1 : 0);
      doy += ICAL.Time.daysInYearPassedMonth[is_leap][12];
      return ICAL.Time.fromDayOfYear(doy, year);
    } else if (doy > ICAL.Time.daysInYearPassedMonth[is_leap][12]) {
      is_leap = (ICAL.Time.isLeapYear(year) ? 1 : 0);
      doy -= ICAL.Time.daysInYearPassedMonth[is_leap][12];
      year++;
      return ICAL.Time.fromDayOfYear(doy, year);
    }

    tt.year = year;
    tt.isDate = true;

    for (var month = 11; month >= 0; month--) {
      if (doy > ICAL.Time.daysInYearPassedMonth[is_leap][month]) {
        tt.month = month + 1;
        tt.day = doy - ICAL.Time.daysInYearPassedMonth[is_leap][month];
        break;
      }
    }

    tt.auto_normalize = true;
    return tt;
  };

  /**
   * Returns a new ICAL.Time instance from a date string, e.g 2015-01-02.
   *
   * @deprecated                Use {@link ICAL.Time.fromDateString} instead
   * @param {String} str        The string to create from
   * @return {ICAL.Time}        The date/time instance
   */
  ICAL.Time.fromStringv2 = function fromString(str) {
    return new ICAL.Time({
      year: parseInt(str.substr(0, 4), 10),
      month: parseInt(str.substr(5, 2), 10),
      day: parseInt(str.substr(8, 2), 10),
      isDate: true
    });
  };

  /**
   * Returns a new ICAL.Time instance from a date string, e.g 2015-01-02.
   *
   * @param {String} aValue     The string to create from
   * @return {ICAL.Time}        The date/time instance
   */
  ICAL.Time.fromDateString = function(aValue) {
    // Dates should have no timezone.
    // Google likes to sometimes specify Z on dates
    // we specifically ignore that to avoid issues.

    // YYYY-MM-DD
    // 2012-10-10
    return new ICAL.Time({
      year: ICAL.helpers.strictParseInt(aValue.substr(0, 4)),
      month: ICAL.helpers.strictParseInt(aValue.substr(5, 2)),
      day: ICAL.helpers.strictParseInt(aValue.substr(8, 2)),
      isDate: true
    });
  };

  /**
   * Returns a new ICAL.Time instance from a date-time string, e.g
   * 2015-01-02T03:04:05. If a property is specified, the timezone is set up
   * from the property's TZID parameter.
   *
   * @param {String} aValue         The string to create from
   * @param {ICAL.Property=} prop   The property the date belongs to
   * @return {ICAL.Time}            The date/time instance
   */
  ICAL.Time.fromDateTimeString = function(aValue, prop) {
    if (aValue.length < 19) {
      throw new Error(
        'invalid date-time value: "' + aValue + '"'
      );
    }

    var zone;

    if (aValue[19] && aValue[19] === 'Z') {
      zone = 'Z';
    } else if (prop) {
      zone = prop.getParameter('tzid');
    }

    // 2012-10-10T10:10:10(Z)?
    var time = new ICAL.Time({
      year: ICAL.helpers.strictParseInt(aValue.substr(0, 4)),
      month: ICAL.helpers.strictParseInt(aValue.substr(5, 2)),
      day: ICAL.helpers.strictParseInt(aValue.substr(8, 2)),
      hour: ICAL.helpers.strictParseInt(aValue.substr(11, 2)),
      minute: ICAL.helpers.strictParseInt(aValue.substr(14, 2)),
      second: ICAL.helpers.strictParseInt(aValue.substr(17, 2)),
      timezone: zone
    });

    return time;
  };

  /**
   * Returns a new ICAL.Time instance from a date or date-time string,
   *
   * @param {String} aValue         The string to create from
   * @param {ICAL.Property=} prop   The property the date belongs to
   * @return {ICAL.Time}            The date/time instance
   */
  ICAL.Time.fromString = function fromString(aValue, aProperty) {
    if (aValue.length > 10) {
      return ICAL.Time.fromDateTimeString(aValue, aProperty);
    } else {
      return ICAL.Time.fromDateString(aValue);
    }
  };

  /**
   * Creates a new ICAL.Time instance from the given Javascript Date.
   *
   * @param {?Date} aDate     The Javascript Date to read, or null to reset
   * @param {Boolean} useUTC  If true, the UTC values of the date will be used
   */
  ICAL.Time.fromJSDate = function fromJSDate(aDate, useUTC) {
    var tt = new ICAL.Time();
    return tt.fromJSDate(aDate, useUTC);
  };

  /**
   * Creates a new ICAL.Time instance from the the passed data object.
   *
   * @param {Object} aData            Time initialization
   * @param {Number=} aData.year      The year for this date
   * @param {Number=} aData.month     The month for this date
   * @param {Number=} aData.day       The day for this date
   * @param {Number=} aData.hour      The hour for this date
   * @param {Number=} aData.minute    The minute for this date
   * @param {Number=} aData.second    The second for this date
   * @param {Boolean=} aData.isDate   If true, the instance represents a date
   *                                    (as opposed to a date-time)
   * @param {ICAL.Timezone=} aZone    Timezone this position occurs in
   */
  ICAL.Time.fromData = function fromData(aData, aZone) {
    var t = new ICAL.Time();
    return t.fromData(aData, aZone);
  };

  /**
   * Creates a new ICAL.Time instance from the current moment.
   * @return {ICAL.Time}
   */
  ICAL.Time.now = function icaltime_now() {
    return ICAL.Time.fromJSDate(new Date(), false);
  };

  /**
   * Returns the date on which ISO week number 1 starts.
   *
   * @see ICAL.Time#weekNumber
   * @param {Number} aYear                  The year to search in
   * @param {ICAL.Time.weekDay=} aWeekStart The week start weekday, used for calculation.
   * @return {ICAL.Time}                    The date on which week number 1 starts
   */
  ICAL.Time.weekOneStarts = function weekOneStarts(aYear, aWeekStart) {
    var t = ICAL.Time.fromData({
      year: aYear,
      month: 1,
      day: 1,
      isDate: true
    });

    var dow = t.dayOfWeek();
    var wkst = aWeekStart || ICAL.Time.DEFAULT_WEEK_START;
    if (dow > ICAL.Time.THURSDAY) {
      t.day += 7;
    }
    if (wkst > ICAL.Time.THURSDAY) {
      t.day -= 7;
    }

    t.day -= dow - wkst;

    return t;
  };

  /**
   * Get the dominical letter for the given year. Letters range from A - G for
   * common years, and AG to GF for leap years.
   *
   * @param {Number} yr           The year to retrieve the letter for
   * @return {String}             The dominical letter.
   */
  ICAL.Time.getDominicalLetter = function(yr) {
    var LTRS = "GFEDCBA";
    var dom = (yr + (yr / 4 | 0) + (yr / 400 | 0) - (yr / 100 | 0) - 1) % 7;
    var isLeap = ICAL.Time.isLeapYear(yr);
    if (isLeap) {
      return LTRS[(dom + 6) % 7] + LTRS[dom];
    } else {
      return LTRS[dom];
    }
  };

  /**
   * January 1st, 1970 as an ICAL.Time.
   * @type {ICAL.Time}
   * @constant
   * @instance
   */
  ICAL.Time.epochTime = ICAL.Time.fromData({
    year: 1970,
    month: 1,
    day: 1,
    hour: 0,
    minute: 0,
    second: 0,
    isDate: false,
    timezone: "Z"
  });

  ICAL.Time._cmp_attr = function _cmp_attr(a, b, attr) {
    if (a[attr] > b[attr]) return 1;
    if (a[attr] < b[attr]) return -1;
    return 0;
  };

  /**
   * The days that have passed in the year after a given month. The array has
   * two members, one being an array of passed days for non-leap years, the
   * other analog for leap years.
   * @example
   * var isLeapYear = ICAL.Time.isLeapYear(year);
   * var passedDays = ICAL.Time.daysInYearPassedMonth[isLeapYear][month];
   * @type {Array.<Array.<Number>>}
   */
  ICAL.Time.daysInYearPassedMonth = [
    [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365],
    [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366]
  ];

  /**
   * The weekday, 1 = SUNDAY, 7 = SATURDAY. Access via
   * ICAL.Time.MONDAY, ICAL.Time.TUESDAY, ...
   *
   * @typedef {Number} weekDay
   * @memberof ICAL.Time
   */

  ICAL.Time.SUNDAY = 1;
  ICAL.Time.MONDAY = 2;
  ICAL.Time.TUESDAY = 3;
  ICAL.Time.WEDNESDAY = 4;
  ICAL.Time.THURSDAY = 5;
  ICAL.Time.FRIDAY = 6;
  ICAL.Time.SATURDAY = 7;

  /**
   * The default weekday for the WKST part.
   * @constant
   * @default ICAL.Time.MONDAY
   */
  ICAL.Time.DEFAULT_WEEK_START = ICAL.Time.MONDAY;
})();
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2015 */



(function() {

  /**
   * Describes a vCard time, which has slight differences to the ICAL.Time.
   * Properties can be null if not specified, for example for dates with
   * reduced accuracy or truncation.
   *
   * Note that currently not all methods are correctly re-implemented for
   * VCardTime. For example, comparison will have undefined results when some
   * members are null.
   *
   * Also, normalization is not yet implemented for this class!
   *
   * @alias ICAL.VCardTime
   * @class
   * @extends {ICAL.Time}
   * @param {Object} data                           The data for the time instance
   * @param {Number=} data.year                     The year for this date
   * @param {Number=} data.month                    The month for this date
   * @param {Number=} data.day                      The day for this date
   * @param {Number=} data.hour                     The hour for this date
   * @param {Number=} data.minute                   The minute for this date
   * @param {Number=} data.second                   The second for this date
   * @param {ICAL.Timezone|ICAL.UtcOffset} zone     The timezone to use
   * @param {String} icaltype                       The type for this date/time object
   */
  ICAL.VCardTime = function(data, zone, icaltype) {
    this.wrappedJSObject = this;
    var time = this._time = Object.create(null);

    time.year = null;
    time.month = null;
    time.day = null;
    time.hour = null;
    time.minute = null;
    time.second = null;

    this.icaltype = icaltype || "date-and-or-time";

    this.fromData(data, zone);
  };
  ICAL.helpers.inherits(ICAL.Time, ICAL.VCardTime, /** @lends ICAL.VCardTime */ {

    /**
     * The class identifier.
     * @constant
     * @type {String}
     * @default "vcardtime"
     */
    icalclass: "vcardtime",

    /**
     * The type name, to be used in the jCal object.
     * @type {String}
     * @default "date-and-or-time"
     */
    icaltype: "date-and-or-time",

    /**
     * The timezone. This can either be floating, UTC, or an instance of
     * ICAL.UtcOffset.
     * @type {ICAL.Timezone|ICAL.UtcOFfset}
     */
    zone: null,

    /**
     * Returns a clone of the vcard date/time object.
     *
     * @return {ICAL.VCardTime}     The cloned object
     */
    clone: function() {
      return new ICAL.VCardTime(this._time, this.zone, this.icaltype);
    },

    _normalize: function() {
      return this;
    },

    /**
     * @inheritdoc
     */
    utcOffset: function() {
      if (this.zone instanceof ICAL.UtcOffset) {
        return this.zone.toSeconds();
      } else {
        return ICAL.Time.prototype.utcOffset.apply(this, arguments);
      }
    },

    /**
     * Returns an RFC 6350 compliant representation of this object.
     *
     * @return {String}         vcard date/time string
     */
    toICALString: function() {
      return ICAL.design.vcard.value[this.icaltype].toICAL(this.toString());
    },

    /**
     * The string representation of this date/time, in jCard form
     * (including : and - separators).
     * @return {String}
     */
    toString: function toString() {
      var p2 = ICAL.helpers.pad2;
      var y = this.year, m = this.month, d = this.day;
      var h = this.hour, mm = this.minute, s = this.second;

      var hasYear = y !== null, hasMonth = m !== null, hasDay = d !== null;
      var hasHour = h !== null, hasMinute = mm !== null, hasSecond = s !== null;

      var datepart = (hasYear ? p2(y) + (hasMonth || hasDay ? '-' : '') : (hasMonth || hasDay ? '--' : '')) +
                     (hasMonth ? p2(m) : '') +
                     (hasDay ? '-' + p2(d) : '');
      var timepart = (hasHour ? p2(h) : '-') + (hasHour && hasMinute ? ':' : '') +
                     (hasMinute ? p2(mm) : '') + (!hasHour && !hasMinute ? '-' : '') +
                     (hasMinute && hasSecond ? ':' : '') +
                     (hasSecond ? p2(s) : '');

      var zone;
      if (this.zone === ICAL.Timezone.utcTimezone) {
        zone = 'Z';
      } else if (this.zone instanceof ICAL.UtcOffset) {
        zone = this.zone.toString();
      } else if (this.zone === ICAL.Timezone.localTimezone) {
        zone = '';
      } else if (this.zone instanceof ICAL.Timezone) {
        var offset = ICAL.UtcOffset.fromSeconds(this.zone.utcOffset(this));
        zone = offset.toString();
      } else {
        zone = '';
      }

      switch (this.icaltype) {
        case "time":
          return timepart + zone;
        case "date-and-or-time":
        case "date-time":
          return datepart + (timepart == '--' ? '' : 'T' + timepart + zone);
        case "date":
          return datepart;
      }
      return null;
    }
  });

  /**
   * Returns a new ICAL.VCardTime instance from a date and/or time string.
   *
   * @param {String} aValue     The string to create from
   * @param {String} aIcalType  The type for this instance, e.g. date-and-or-time
   * @return {ICAL.VCardTime}   The date/time instance
   */
  ICAL.VCardTime.fromDateAndOrTimeString = function(aValue, aIcalType) {
    function part(v, s, e) {
      return v ? ICAL.helpers.strictParseInt(v.substr(s, e)) : null;
    }
    var parts = aValue.split('T');
    var dt = parts[0], tmz = parts[1];
    var splitzone = tmz ? ICAL.design.vcard.value.time._splitZone(tmz) : [];
    var zone = splitzone[0], tm = splitzone[1];

    var stoi = ICAL.helpers.strictParseInt;
    var dtlen = dt ? dt.length : 0;
    var tmlen = tm ? tm.length : 0;

    var hasDashDate = dt && dt[0] == '-' && dt[1] == '-';
    var hasDashTime = tm && tm[0] == '-';

    var o = {
      year: hasDashDate ? null : part(dt, 0, 4),
      month: hasDashDate && (dtlen == 4 || dtlen == 7) ? part(dt, 2, 2) : dtlen == 7 ? part(dt, 5, 2) : dtlen == 10 ? part(dt, 5, 2) : null,
      day: dtlen == 5 ? part(dt, 3, 2) : dtlen == 7 && hasDashDate ? part(dt, 5, 2) : dtlen == 10 ? part(dt, 8, 2) : null,

      hour: hasDashTime ? null : part(tm, 0, 2),
      minute: hasDashTime && tmlen == 3 ? part(tm, 1, 2) : tmlen > 4 ? hasDashTime ? part(tm, 1, 2) : part(tm, 3, 2) : null,
      second: tmlen == 4 ? part(tm, 2, 2) : tmlen == 6 ? part(tm, 4, 2) : tmlen == 8 ? part(tm, 6, 2) : null
    };

    if (zone == 'Z') {
      zone = ICAL.Timezone.utcTimezone;
    } else if (zone && zone[3] == ':') {
      zone = ICAL.UtcOffset.fromString(zone);
    } else {
      zone = null;
    }

    return new ICAL.VCardTime(o, zone, aIcalType);
  };
})();
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */



(function() {
  var DOW_MAP = {
    SU: ICAL.Time.SUNDAY,
    MO: ICAL.Time.MONDAY,
    TU: ICAL.Time.TUESDAY,
    WE: ICAL.Time.WEDNESDAY,
    TH: ICAL.Time.THURSDAY,
    FR: ICAL.Time.FRIDAY,
    SA: ICAL.Time.SATURDAY
  };

  var REVERSE_DOW_MAP = {};
  for (var key in DOW_MAP) {
    /* istanbul ignore else */
    if (DOW_MAP.hasOwnProperty(key)) {
      REVERSE_DOW_MAP[DOW_MAP[key]] = key;
    }
  }

  var COPY_PARTS = ["BYSECOND", "BYMINUTE", "BYHOUR", "BYDAY",
                    "BYMONTHDAY", "BYYEARDAY", "BYWEEKNO",
                    "BYMONTH", "BYSETPOS"];

  /**
   * @classdesc
   * This class represents the "recur" value type, with various calculation
   * and manipulation methods.
   *
   * @class
   * @alias ICAL.Recur
   * @param {Object} data                               An object with members of the recurrence
   * @param {ICAL.Recur.frequencyValues=} data.freq     The frequency value
   * @param {Number=} data.interval                     The INTERVAL value
   * @param {ICAL.Time.weekDay=} data.wkst              The week start value
   * @param {ICAL.Time=} data.until                     The end of the recurrence set
   * @param {Number=} data.count                        The number of occurrences
   * @param {Array.<Number>=} data.bysecond             The seconds for the BYSECOND part
   * @param {Array.<Number>=} data.byminute             The minutes for the BYMINUTE part
   * @param {Array.<Number>=} data.byhour               The hours for the BYHOUR part
   * @param {Array.<String>=} data.byday                The BYDAY values
   * @param {Array.<Number>=} data.bymonthday           The days for the BYMONTHDAY part
   * @param {Array.<Number>=} data.byyearday            The days for the BYYEARDAY part
   * @param {Array.<Number>=} data.byweekno             The weeks for the BYWEEKNO part
   * @param {Array.<Number>=} data.bymonth              The month for the BYMONTH part
   * @param {Array.<Number>=} data.bysetpos             The positionals for the BYSETPOS part
   */
  ICAL.Recur = function icalrecur(data) {
    this.wrappedJSObject = this;
    this.parts = {};

    if (data && typeof(data) === 'object') {
      this.fromData(data);
    }
  };

  ICAL.Recur.prototype = {
    /**
     * An object holding the BY-parts of the recurrence rule
     * @type {Object}
     */
    parts: null,

    /**
     * The interval value for the recurrence rule.
     * @type {Number}
     */
    interval: 1,

    /**
     * The week start day
     *
     * @type {ICAL.Time.weekDay}
     * @default ICAL.Time.MONDAY
     */
    wkst: ICAL.Time.MONDAY,

    /**
     * The end of the recurrence
     * @type {?ICAL.Time}
     */
    until: null,

    /**
     * The maximum number of occurrences
     * @type {?Number}
     */
    count: null,

    /**
     * The frequency value.
     * @type {ICAL.Recur.frequencyValues}
     */
    freq: null,

    /**
     * The class identifier.
     * @constant
     * @type {String}
     * @default "icalrecur"
     */
    icalclass: "icalrecur",

    /**
     * The type name, to be used in the jCal object.
     * @constant
     * @type {String}
     * @default "recur"
     */
    icaltype: "recur",

    /**
     * Create a new iterator for this recurrence rule. The passed start date
     * must be the start date of the event, not the start of the range to
     * search in.
     *
     * @example
     * var recur = comp.getFirstPropertyValue('rrule');
     * var dtstart = comp.getFirstPropertyValue('dtstart');
     * var iter = recur.iterator(dtstart);
     * for (var next = iter.next(); next; next = iter.next()) {
     *   if (next.compare(rangeStart) < 0) {
     *     continue;
     *   }
     *   console.log(next.toString());
     * }
     *
     * @param {ICAL.Time} aStart        The item's start date
     * @return {ICAL.RecurIterator}     The recurrence iterator
     */
    iterator: function(aStart) {
      return new ICAL.RecurIterator({
        rule: this,
        dtstart: aStart
      });
    },

    /**
     * Returns a clone of the recurrence object.
     *
     * @return {ICAL.Recur}      The cloned object
     */
    clone: function clone() {
      return new ICAL.Recur(this.toJSON());
    },

    /**
     * Checks if the current rule is finite, i.e. has a count or until part.
     *
     * @return {Boolean}        True, if the rule is finite
     */
    isFinite: function isfinite() {
      return !!(this.count || this.until);
    },

    /**
     * Checks if the current rule has a count part, and not limited by an until
     * part.
     *
     * @return {Boolean}        True, if the rule is by count
     */
    isByCount: function isbycount() {
      return !!(this.count && !this.until);
    },

    /**
     * Adds a component (part) to the recurrence rule. This is not a component
     * in the sense of {@link ICAL.Component}, but a part of the recurrence
     * rule, i.e. BYMONTH.
     *
     * @param {String} aType            The name of the component part
     * @param {Array|String} aValue     The component value
     */
    addComponent: function addPart(aType, aValue) {
      var ucname = aType.toUpperCase();
      if (ucname in this.parts) {
        this.parts[ucname].push(aValue);
      } else {
        this.parts[ucname] = [aValue];
      }
    },

    /**
     * Sets the component value for the given by-part.
     *
     * @param {String} aType        The component part name
     * @param {Array} aValues       The component values
     */
    setComponent: function setComponent(aType, aValues) {
      this.parts[aType.toUpperCase()] = aValues.slice();
    },

    /**
     * Gets (a copy) of the requested component value.
     *
     * @param {String} aType        The component part name
     * @return {Array}              The component part value
     */
    getComponent: function getComponent(aType) {
      var ucname = aType.toUpperCase();
      return (ucname in this.parts ? this.parts[ucname].slice() : []);
    },

    /**
     * Retrieves the next occurrence after the given recurrence id. See the
     * guide on {@tutorial terminology} for more details.
     *
     * NOTE: Currently, this method iterates all occurrences from the start
     * date. It should not be called in a loop for performance reasons. If you
     * would like to get more than one occurrence, you can iterate the
     * occurrences manually, see the example on the
     * {@link ICAL.Recur#iterator iterator} method.
     *
     * @param {ICAL.Time} aStartTime        The start of the event series
     * @param {ICAL.Time} aRecurrenceId     The date of the last occurrence
     * @return {ICAL.Time}                  The next occurrence after
     */
    getNextOccurrence: function getNextOccurrence(aStartTime, aRecurrenceId) {
      var iter = this.iterator(aStartTime);
      var next, cdt;

      do {
        next = iter.next();
      } while (next && next.compare(aRecurrenceId) <= 0);

      if (next && aRecurrenceId.zone) {
        next.zone = aRecurrenceId.zone;
      }

      return next;
    },

    /**
     * Sets up the current instance using members from the passed data object.
     *
     * @param {Object} data                               An object with members of the recurrence
     * @param {ICAL.Recur.frequencyValues=} data.freq     The frequency value
     * @param {Number=} data.interval                     The INTERVAL value
     * @param {ICAL.Time.weekDay=} data.wkst              The week start value
     * @param {ICAL.Time=} data.until                     The end of the recurrence set
     * @param {Number=} data.count                        The number of occurrences
     * @param {Array.<Number>=} data.bysecond             The seconds for the BYSECOND part
     * @param {Array.<Number>=} data.byminute             The minutes for the BYMINUTE part
     * @param {Array.<Number>=} data.byhour               The hours for the BYHOUR part
     * @param {Array.<String>=} data.byday                The BYDAY values
     * @param {Array.<Number>=} data.bymonthday           The days for the BYMONTHDAY part
     * @param {Array.<Number>=} data.byyearday            The days for the BYYEARDAY part
     * @param {Array.<Number>=} data.byweekno             The weeks for the BYWEEKNO part
     * @param {Array.<Number>=} data.bymonth              The month for the BYMONTH part
     * @param {Array.<Number>=} data.bysetpos             The positionals for the BYSETPOS part
     */
    fromData: function(data) {
      for (var key in data) {
        var uckey = key.toUpperCase();

        if (uckey in partDesign) {
          if (Array.isArray(data[key])) {
            this.parts[uckey] = data[key];
          } else {
            this.parts[uckey] = [data[key]];
          }
        } else {
          this[key] = data[key];
        }
      }

      if (this.interval && typeof this.interval != "number") {
        optionDesign.INTERVAL(this.interval, this);
      }

      if (this.wkst && typeof this.wkst != "number") {
        this.wkst = ICAL.Recur.icalDayToNumericDay(this.wkst);
      }

      if (this.until && !(this.until instanceof ICAL.Time)) {
        this.until = ICAL.Time.fromString(this.until);
      }
    },

    /**
     * The jCal representation of this recurrence type.
     * @return {Object}
     */
    toJSON: function() {
      var res = Object.create(null);
      res.freq = this.freq;

      if (this.count) {
        res.count = this.count;
      }

      if (this.interval > 1) {
        res.interval = this.interval;
      }

      for (var k in this.parts) {
        /* istanbul ignore if */
        if (!this.parts.hasOwnProperty(k)) {
          continue;
        }
        var kparts = this.parts[k];
        if (Array.isArray(kparts) && kparts.length == 1) {
          res[k.toLowerCase()] = kparts[0];
        } else {
          res[k.toLowerCase()] = ICAL.helpers.clone(this.parts[k]);
        }
      }

      if (this.until) {
        res.until = this.until.toString();
      }
      if ('wkst' in this && this.wkst !== ICAL.Time.DEFAULT_WEEK_START) {
        res.wkst = ICAL.Recur.numericDayToIcalDay(this.wkst);
      }
      return res;
    },

    /**
     * The string representation of this recurrence rule.
     * @return {String}
     */
    toString: function icalrecur_toString() {
      // TODO retain order
      var str = "FREQ=" + this.freq;
      if (this.count) {
        str += ";COUNT=" + this.count;
      }
      if (this.interval > 1) {
        str += ";INTERVAL=" + this.interval;
      }
      for (var k in this.parts) {
        /* istanbul ignore else */
        if (this.parts.hasOwnProperty(k)) {
          str += ";" + k + "=" + this.parts[k];
        }
      }
      if (this.until) {
        str += ';UNTIL=' + this.until.toICALString();
      }
      if ('wkst' in this && this.wkst !== ICAL.Time.DEFAULT_WEEK_START) {
        str += ';WKST=' + ICAL.Recur.numericDayToIcalDay(this.wkst);
      }
      return str;
    }
  };

  function parseNumericValue(type, min, max, value) {
    var result = value;

    if (value[0] === '+') {
      result = value.substr(1);
    }

    result = ICAL.helpers.strictParseInt(result);

    if (min !== undefined && value < min) {
      throw new Error(
        type + ': invalid value "' + value + '" must be > ' + min
      );
    }

    if (max !== undefined && value > max) {
      throw new Error(
        type + ': invalid value "' + value + '" must be < ' + min
      );
    }

    return result;
  }

  /**
   * Convert an ical representation of a day (SU, MO, etc..)
   * into a numeric value of that day.
   *
   * @param {String} string     The iCalendar day name
   * @param {ICAL.Time.weekDay=} aWeekStart
   *        The week start weekday, defaults to SUNDAY
   * @return {Number}           Numeric value of given day
   */
  ICAL.Recur.icalDayToNumericDay = function toNumericDay(string, aWeekStart) {
    //XXX: this is here so we can deal
    //     with possibly invalid string values.
    var firstDow = aWeekStart || ICAL.Time.SUNDAY;
    return ((DOW_MAP[string] - firstDow + 7) % 7) + 1;
  };

  /**
   * Convert a numeric day value into its ical representation (SU, MO, etc..)
   *
   * @param {Number} num        Numeric value of given day
   * @param {ICAL.Time.weekDay=} aWeekStart
   *        The week start weekday, defaults to SUNDAY
   * @return {String}           The ICAL day value, e.g SU,MO,...
   */
  ICAL.Recur.numericDayToIcalDay = function toIcalDay(num, aWeekStart) {
    //XXX: this is here so we can deal with possibly invalid number values.
    //     Also, this allows consistent mapping between day numbers and day
    //     names for external users.
    var firstDow = aWeekStart || ICAL.Time.SUNDAY;
    var dow = (num + firstDow - ICAL.Time.SUNDAY);
    if (dow > 7) {
      dow -= 7;
    }
    return REVERSE_DOW_MAP[dow];
  };

  var VALID_DAY_NAMES = /^(SU|MO|TU|WE|TH|FR|SA)$/;
  var VALID_BYDAY_PART = /^([+-])?(5[0-3]|[1-4][0-9]|[1-9])?(SU|MO|TU|WE|TH|FR|SA)$/;

  /**
   * Possible frequency values for the FREQ part
   * (YEARLY, MONTHLY, WEEKLY, DAILY, HOURLY, MINUTELY, SECONDLY)
   *
   * @typedef {String} frequencyValues
   * @memberof ICAL.Recur
   */

  var ALLOWED_FREQ = ['SECONDLY', 'MINUTELY', 'HOURLY',
                      'DAILY', 'WEEKLY', 'MONTHLY', 'YEARLY'];

  var optionDesign = {
    FREQ: function(value, dict, fmtIcal) {
      // yes this is actually equal or faster then regex.
      // upside here is we can enumerate the valid values.
      if (ALLOWED_FREQ.indexOf(value) !== -1) {
        dict.freq = value;
      } else {
        throw new Error(
          'invalid frequency "' + value + '" expected: "' +
          ALLOWED_FREQ.join(', ') + '"'
        );
      }
    },

    COUNT: function(value, dict, fmtIcal) {
      dict.count = ICAL.helpers.strictParseInt(value);
    },

    INTERVAL: function(value, dict, fmtIcal) {
      dict.interval = ICAL.helpers.strictParseInt(value);
      if (dict.interval < 1) {
        // 0 or negative values are not allowed, some engines seem to generate
        // it though. Assume 1 instead.
        dict.interval = 1;
      }
    },

    UNTIL: function(value, dict, fmtIcal) {
      if (value.length > 10) {
        dict.until = ICAL.design.icalendar.value['date-time'].fromICAL(value);
      } else {
        dict.until = ICAL.design.icalendar.value.date.fromICAL(value);
      }
      if (!fmtIcal) {
        dict.until = ICAL.Time.fromString(dict.until);
      }
    },

    WKST: function(value, dict, fmtIcal) {
      if (VALID_DAY_NAMES.test(value)) {
        dict.wkst = ICAL.Recur.icalDayToNumericDay(value);
      } else {
        throw new Error('invalid WKST value "' + value + '"');
      }
    }
  };

  var partDesign = {
    BYSECOND: parseNumericValue.bind(this, 'BYSECOND', 0, 60),
    BYMINUTE: parseNumericValue.bind(this, 'BYMINUTE', 0, 59),
    BYHOUR: parseNumericValue.bind(this, 'BYHOUR', 0, 23),
    BYDAY: function(value) {
      if (VALID_BYDAY_PART.test(value)) {
        return value;
      } else {
        throw new Error('invalid BYDAY value "' + value + '"');
      }
    },
    BYMONTHDAY: parseNumericValue.bind(this, 'BYMONTHDAY', -31, 31),
    BYYEARDAY: parseNumericValue.bind(this, 'BYYEARDAY', -366, 366),
    BYWEEKNO: parseNumericValue.bind(this, 'BYWEEKNO', -53, 53),
    BYMONTH: parseNumericValue.bind(this, 'BYMONTH', 0, 12),
    BYSETPOS: parseNumericValue.bind(this, 'BYSETPOS', -366, 366)
  };


  /**
   * Creates a new {@link ICAL.Recur} instance from the passed string.
   *
   * @param {String} string         The string to parse
   * @return {ICAL.Recur}           The created recurrence instance
   */
  ICAL.Recur.fromString = function(string) {
    var data = ICAL.Recur._stringToData(string, false);
    return new ICAL.Recur(data);
  };

  /**
   * Creates a new {@link ICAL.Recur} instance using members from the passed
   * data object.
   *
   * @param {Object} aData                              An object with members of the recurrence
   * @param {ICAL.Recur.frequencyValues=} aData.freq    The frequency value
   * @param {Number=} aData.interval                    The INTERVAL value
   * @param {ICAL.Time.weekDay=} aData.wkst             The week start value
   * @param {ICAL.Time=} aData.until                    The end of the recurrence set
   * @param {Number=} aData.count                       The number of occurrences
   * @param {Array.<Number>=} aData.bysecond            The seconds for the BYSECOND part
   * @param {Array.<Number>=} aData.byminute            The minutes for the BYMINUTE part
   * @param {Array.<Number>=} aData.byhour              The hours for the BYHOUR part
   * @param {Array.<String>=} aData.byday               The BYDAY values
   * @param {Array.<Number>=} aData.bymonthday          The days for the BYMONTHDAY part
   * @param {Array.<Number>=} aData.byyearday           The days for the BYYEARDAY part
   * @param {Array.<Number>=} aData.byweekno            The weeks for the BYWEEKNO part
   * @param {Array.<Number>=} aData.bymonth             The month for the BYMONTH part
   * @param {Array.<Number>=} aData.bysetpos            The positionals for the BYSETPOS part
   */
  ICAL.Recur.fromData = function(aData) {
    return new ICAL.Recur(aData);
  };

  /**
   * Converts a recurrence string to a data object, suitable for the fromData
   * method.
   *
   * @param {String} string     The string to parse
   * @param {Boolean} fmtIcal   If true, the string is considered to be an
   *                              iCalendar string
   * @return {ICAL.Recur}       The recurrence instance
   */
  ICAL.Recur._stringToData = function(string, fmtIcal) {
    var dict = Object.create(null);

    // split is slower in FF but fast enough.
    // v8 however this is faster then manual split?
    var values = string.split(';');
    var len = values.length;

    for (var i = 0; i < len; i++) {
      var parts = values[i].split('=');
      var ucname = parts[0].toUpperCase();
      var lcname = parts[0].toLowerCase();
      var name = (fmtIcal ? lcname : ucname);
      var value = parts[1];

      if (ucname in partDesign) {
        var partArr = value.split(',');
        var partArrIdx = 0;
        var partArrLen = partArr.length;

        for (; partArrIdx < partArrLen; partArrIdx++) {
          partArr[partArrIdx] = partDesign[ucname](partArr[partArrIdx]);
        }
        dict[name] = (partArr.length == 1 ? partArr[0] : partArr);
      } else if (ucname in optionDesign) {
        optionDesign[ucname](value, dict, fmtIcal);
      } else {
        // Don't swallow unknown values. Just set them as they are.
        dict[lcname] = value;
      }
    }

    return dict;
  };
})();
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.RecurIterator = (function() {

  /**
   * @classdesc
   * An iterator for a single recurrence rule. This class usually doesn't have
   * to be instanciated directly, the convenience method
   * {@link ICAL.Recur#iterator} can be used.
   *
   * @description
   * The options object may contain additional members when resuming iteration from a previous run
   *
   * @description
   * The options object may contain additional members when resuming iteration
   * from a previous run.
   *
   * @class
   * @alias ICAL.RecurIterator
   * @param {Object} options                The iterator options
   * @param {ICAL.Recur} options.rule       The rule to iterate.
   * @param {ICAL.Time} options.dtstart     The start date of the event.
   * @param {Boolean=} options.initialized  When true, assume that options are
   *        from a previously constructed iterator. Initialization will not be
   *        repeated.
   */
  function icalrecur_iterator(options) {
    this.fromData(options);
  }

  icalrecur_iterator.prototype = {

    /**
     * True when iteration is finished.
     * @type {Boolean}
     */
    completed: false,

    /**
     * The rule that is being iterated
     * @type {ICAL.Recur}
     */
    rule: null,

    /**
     * The start date of the event being iterated.
     * @type {ICAL.Time}
     */
    dtstart: null,

    /**
     * The last occurrence that was returned from the
     * {@link ICAL.RecurIterator#next} method.
     * @type {ICAL.Time}
     */
    last: null,

    /**
     * The sequence number from the occurrence
     * @type {Number}
     */
    occurrence_number: 0,

    /**
     * The indices used for the {@link ICAL.RecurIterator#by_data} object.
     * @type {Object}
     * @private
     */
    by_indices: null,

    /**
     * If true, the iterator has already been initialized
     * @type {Boolean}
     * @private
     */
    initialized: false,

    /**
     * The initializd by-data.
     * @type {Object}
     * @private
     */
    by_data: null,

    /**
     * The expanded yeardays
     * @type {Array}
     * @private
     */
    days: null,

    /**
     * The index in the {@link ICAL.RecurIterator#days} array.
     * @type {Number}
     * @private
     */
    days_index: 0,

    /**
     * Initialize the recurrence iterator from the passed data object. This
     * method is usually not called directly, you can initialize the iterator
     * through the constructor.
     *
     * @param {Object} options                The iterator options
     * @param {ICAL.Recur} options.rule       The rule to iterate.
     * @param {ICAL.Time} options.dtstart     The start date of the event.
     * @param {Boolean=} options.initialized  When true, assume that options are
     *        from a previously constructed iterator. Initialization will not be
     *        repeated.
     */
    fromData: function(options) {
      this.rule = ICAL.helpers.formatClassType(options.rule, ICAL.Recur);

      if (!this.rule) {
        throw new Error('iterator requires a (ICAL.Recur) rule');
      }

      this.dtstart = ICAL.helpers.formatClassType(options.dtstart, ICAL.Time);

      if (!this.dtstart) {
        throw new Error('iterator requires a (ICAL.Time) dtstart');
      }

      if (options.by_data) {
        this.by_data = options.by_data;
      } else {
        this.by_data = ICAL.helpers.clone(this.rule.parts, true);
      }

      if (options.occurrence_number)
        this.occurrence_number = options.occurrence_number;

      this.days = options.days || [];
      if (options.last) {
        this.last = ICAL.helpers.formatClassType(options.last, ICAL.Time);
      }

      this.by_indices = options.by_indices;

      if (!this.by_indices) {
        this.by_indices = {
          "BYSECOND": 0,
          "BYMINUTE": 0,
          "BYHOUR": 0,
          "BYDAY": 0,
          "BYMONTH": 0,
          "BYWEEKNO": 0,
          "BYMONTHDAY": 0
        };
      }

      this.initialized = options.initialized || false;

      if (!this.initialized) {
        this.init();
      }
    },

    /**
     * Intialize the iterator
     * @private
     */
    init: function icalrecur_iterator_init() {
      this.initialized = true;
      this.last = this.dtstart.clone();
      var parts = this.by_data;

      if ("BYDAY" in parts) {
        // libical does this earlier when the rule is loaded, but we postpone to
        // now so we can preserve the original order.
        this.sort_byday_rules(parts.BYDAY);
      }

      // If the BYYEARDAY appares, no other date rule part may appear
      if ("BYYEARDAY" in parts) {
        if ("BYMONTH" in parts || "BYWEEKNO" in parts ||
            "BYMONTHDAY" in parts || "BYDAY" in parts) {
          throw new Error("Invalid BYYEARDAY rule");
        }
      }

      // BYWEEKNO and BYMONTHDAY rule parts may not both appear
      if ("BYWEEKNO" in parts && "BYMONTHDAY" in parts) {
        throw new Error("BYWEEKNO does not fit to BYMONTHDAY");
      }

      // For MONTHLY recurrences (FREQ=MONTHLY) neither BYYEARDAY nor
      // BYWEEKNO may appear.
      if (this.rule.freq == "MONTHLY" &&
          ("BYYEARDAY" in parts || "BYWEEKNO" in parts)) {
        throw new Error("For MONTHLY recurrences neither BYYEARDAY nor BYWEEKNO may appear");
      }

      // For WEEKLY recurrences (FREQ=WEEKLY) neither BYMONTHDAY nor
      // BYYEARDAY may appear.
      if (this.rule.freq == "WEEKLY" &&
          ("BYYEARDAY" in parts || "BYMONTHDAY" in parts)) {
        throw new Error("For WEEKLY recurrences neither BYMONTHDAY nor BYYEARDAY may appear");
      }

      // BYYEARDAY may only appear in YEARLY rules
      if (this.rule.freq != "YEARLY" && "BYYEARDAY" in parts) {
        throw new Error("BYYEARDAY may only appear in YEARLY rules");
      }

      this.last.second = this.setup_defaults("BYSECOND", "SECONDLY", this.dtstart.second);
      this.last.minute = this.setup_defaults("BYMINUTE", "MINUTELY", this.dtstart.minute);
      this.last.hour = this.setup_defaults("BYHOUR", "HOURLY", this.dtstart.hour);
      this.last.day = this.setup_defaults("BYMONTHDAY", "DAILY", this.dtstart.day);
      this.last.month = this.setup_defaults("BYMONTH", "MONTHLY", this.dtstart.month);

      if (this.rule.freq == "WEEKLY") {
        if ("BYDAY" in parts) {
          var bydayParts = this.ruleDayOfWeek(parts.BYDAY[0], this.rule.wkst);
          var pos = bydayParts[0];
          var dow = bydayParts[1];
          var wkdy = dow - this.last.dayOfWeek(this.rule.wkst);
          if ((this.last.dayOfWeek(this.rule.wkst) < dow && wkdy >= 0) || wkdy < 0) {
            // Initial time is after first day of BYDAY data
            this.last.day += wkdy;
          }
        } else {
          var dayName = ICAL.Recur.numericDayToIcalDay(this.dtstart.dayOfWeek());
          parts.BYDAY = [dayName];
        }
      }

      if (this.rule.freq == "YEARLY") {
        for (;;) {
          this.expand_year_days(this.last.year);
          if (this.days.length > 0) {
            break;
          }
          this.increment_year(this.rule.interval);
        }

        this._nextByYearDay();
      }

      if (this.rule.freq == "MONTHLY" && this.has_by_data("BYDAY")) {
        var tempLast = null;
        var initLast = this.last.clone();
        var daysInMonth = ICAL.Time.daysInMonth(this.last.month, this.last.year);

        // Check every weekday in BYDAY with relative dow and pos.
        for (var i in this.by_data.BYDAY) {
          /* istanbul ignore if */
          if (!this.by_data.BYDAY.hasOwnProperty(i)) {
            continue;
          }
          this.last = initLast.clone();
          var bydayParts = this.ruleDayOfWeek(this.by_data.BYDAY[i]);
          var pos = bydayParts[0];
          var dow = bydayParts[1];
          var dayOfMonth = this.last.nthWeekDay(dow, pos);

          // If |pos| >= 6, the byday is invalid for a monthly rule.
          if (pos >= 6 || pos <= -6) {
            throw new Error("Malformed values in BYDAY part");
          }

          // If a Byday with pos=+/-5 is not in the current month it
          // must be searched in the next months.
          if (dayOfMonth > daysInMonth || dayOfMonth <= 0) {
            // Skip if we have already found a "last" in this month.
            if (tempLast && tempLast.month == initLast.month) {
              continue;
            }
            while (dayOfMonth > daysInMonth || dayOfMonth <= 0) {
              this.increment_month();
              daysInMonth = ICAL.Time.daysInMonth(this.last.month, this.last.year);
              dayOfMonth = this.last.nthWeekDay(dow, pos);
            }
          }

          this.last.day = dayOfMonth;
          if (!tempLast || this.last.compare(tempLast) < 0) {
            tempLast = this.last.clone();
          }
        }
        this.last = tempLast.clone();

        //XXX: This feels like a hack, but we need to initialize
        //     the BYMONTHDAY case correctly and byDayAndMonthDay handles
        //     this case. It accepts a special flag which will avoid incrementing
        //     the initial value without the flag days that match the start time
        //     would be missed.
        if (this.has_by_data('BYMONTHDAY')) {
          this._byDayAndMonthDay(true);
        }

        if (this.last.day > daysInMonth || this.last.day == 0) {
          throw new Error("Malformed values in BYDAY part");
        }

      } else if (this.has_by_data("BYMONTHDAY")) {
        if (this.last.day < 0) {
          var daysInMonth = ICAL.Time.daysInMonth(this.last.month, this.last.year);
          this.last.day = daysInMonth + this.last.day + 1;
        }
      }

    },

    /**
     * Retrieve the next occurrence from the iterator.
     * @return {ICAL.Time}
     */
    next: function icalrecur_iterator_next() {
      var before = (this.last ? this.last.clone() : null);

      if ((this.rule.count && this.occurrence_number >= this.rule.count) ||
          (this.rule.until && this.last.compare(this.rule.until) > 0)) {

        //XXX: right now this is just a flag and has no impact
        //     we can simplify the above case to check for completed later.
        this.completed = true;

        return null;
      }

      if (this.occurrence_number == 0 && this.last.compare(this.dtstart) >= 0) {
        // First of all, give the instance that was initialized
        this.occurrence_number++;
        return this.last;
      }


      var valid;
      do {
        valid = 1;

        switch (this.rule.freq) {
        case "SECONDLY":
          this.next_second();
          break;
        case "MINUTELY":
          this.next_minute();
          break;
        case "HOURLY":
          this.next_hour();
          break;
        case "DAILY":
          this.next_day();
          break;
        case "WEEKLY":
          this.next_week();
          break;
        case "MONTHLY":
          valid = this.next_month();
          break;
        case "YEARLY":
          this.next_year();
          break;

        default:
          return null;
        }
      } while (!this.check_contracting_rules() ||
               this.last.compare(this.dtstart) < 0 ||
               !valid);

      // TODO is this valid?
      if (this.last.compare(before) == 0) {
        throw new Error("Same occurrence found twice, protecting " +
                        "you from death by recursion");
      }

      if (this.rule.until && this.last.compare(this.rule.until) > 0) {
        this.completed = true;
        return null;
      } else {
        this.occurrence_number++;
        return this.last;
      }
    },

    next_second: function next_second() {
      return this.next_generic("BYSECOND", "SECONDLY", "second", "minute");
    },

    increment_second: function increment_second(inc) {
      return this.increment_generic(inc, "second", 60, "minute");
    },

    next_minute: function next_minute() {
      return this.next_generic("BYMINUTE", "MINUTELY",
                               "minute", "hour", "next_second");
    },

    increment_minute: function increment_minute(inc) {
      return this.increment_generic(inc, "minute", 60, "hour");
    },

    next_hour: function next_hour() {
      return this.next_generic("BYHOUR", "HOURLY", "hour",
                               "monthday", "next_minute");
    },

    increment_hour: function increment_hour(inc) {
      this.increment_generic(inc, "hour", 24, "monthday");
    },

    next_day: function next_day() {
      var has_by_day = ("BYDAY" in this.by_data);
      var this_freq = (this.rule.freq == "DAILY");

      if (this.next_hour() == 0) {
        return 0;
      }

      if (this_freq) {
        this.increment_monthday(this.rule.interval);
      } else {
        this.increment_monthday(1);
      }

      return 0;
    },

    next_week: function next_week() {
      var end_of_data = 0;

      if (this.next_weekday_by_week() == 0) {
        return end_of_data;
      }

      if (this.has_by_data("BYWEEKNO")) {
        var idx = ++this.by_indices.BYWEEKNO;

        if (this.by_indices.BYWEEKNO == this.by_data.BYWEEKNO.length) {
          this.by_indices.BYWEEKNO = 0;
          end_of_data = 1;
        }

        // HACK should be first month of the year
        this.last.month = 1;
        this.last.day = 1;

        var week_no = this.by_data.BYWEEKNO[this.by_indices.BYWEEKNO];

        this.last.day += 7 * week_no;

        if (end_of_data) {
          this.increment_year(1);
        }
      } else {
        // Jump to the next week
        this.increment_monthday(7 * this.rule.interval);
      }

      return end_of_data;
    },

    /**
     * Normalize each by day rule for a given year/month.
     * Takes into account ordering and negative rules
     *
     * @private
     * @param {Number} year         Current year.
     * @param {Number} month        Current month.
     * @param {Array}  rules        Array of rules.
     *
     * @return {Array} sorted and normalized rules.
     *                 Negative rules will be expanded to their
     *                 correct positive values for easier processing.
     */
    normalizeByMonthDayRules: function(year, month, rules) {
      var daysInMonth = ICAL.Time.daysInMonth(month, year);

      // XXX: This is probably bad for performance to allocate
      //      a new array for each month we scan, if possible
      //      we should try to optimize this...
      var newRules = [];

      var ruleIdx = 0;
      var len = rules.length;
      var rule;

      for (; ruleIdx < len; ruleIdx++) {
        rule = rules[ruleIdx];

        // if this rule falls outside of given
        // month discard it.
        if (Math.abs(rule) > daysInMonth) {
          continue;
        }

        // negative case
        if (rule < 0) {
          // we add (not subtract its a negative number)
          // one from the rule because 1 === last day of month
          rule = daysInMonth + (rule + 1);
        } else if (rule === 0) {
          // skip zero its invalid.
          continue;
        }

        // only add unique items...
        if (newRules.indexOf(rule) === -1) {
          newRules.push(rule);
        }

      }

      // unique and sort
      return newRules.sort(function(a, b) { return a - b; });
    },

    /**
     * NOTES:
     * We are given a list of dates in the month (BYMONTHDAY) (23, etc..)
     * Also we are given a list of days (BYDAY) (MO, 2SU, etc..) when
     * both conditions match a given date (this.last.day) iteration stops.
     *
     * @private
     * @param {Boolean=} isInit     When given true will not increment the
     *                                current day (this.last).
     */
    _byDayAndMonthDay: function(isInit) {
      var byMonthDay; // setup in initMonth
      var byDay = this.by_data.BYDAY;

      var date;
      var dateIdx = 0;
      var dateLen; // setup in initMonth
      var dayLen = byDay.length;

      // we are not valid by default
      var dataIsValid = 0;

      var daysInMonth;
      var self = this;
      // we need a copy of this, because a DateTime gets normalized
      // automatically if the day is out of range. At some points we
      // set the last day to 0 to start counting.
      var lastDay = this.last.day;

      function initMonth() {
        daysInMonth = ICAL.Time.daysInMonth(
          self.last.month, self.last.year
        );

        byMonthDay = self.normalizeByMonthDayRules(
          self.last.year,
          self.last.month,
          self.by_data.BYMONTHDAY
        );

        dateLen = byMonthDay.length;

        // For the case of more than one occurrence in one month
        // we have to be sure to start searching after the last
        // found date or at the last BYMONTHDAY, unless we are
        // initializing the iterator because in this case we have
        // to consider the last found date too.
        while (byMonthDay[dateIdx] <= lastDay &&
               !(isInit && byMonthDay[dateIdx] == lastDay) &&
               dateIdx < dateLen - 1) {
          dateIdx++;
        }
      }

      function nextMonth() {
        // since the day is incremented at the start
        // of the loop below, we need to start at 0
        lastDay = 0;
        self.increment_month();
        dateIdx = 0;
        initMonth();
      }

      initMonth();

      // should come after initMonth
      if (isInit) {
        lastDay -= 1;
      }

      // Use a counter to avoid an infinite loop with malformed rules.
      // Stop checking after 4 years so we consider also a leap year.
      var monthsCounter = 48;

      while (!dataIsValid && monthsCounter) {
        monthsCounter--;
        // increment the current date. This is really
        // important otherwise we may fall into the infinite
        // loop trap. The initial date takes care of the case
        // where the current date is the date we are looking
        // for.
        date = lastDay + 1;

        if (date > daysInMonth) {
          nextMonth();
          continue;
        }

        // find next date
        var next = byMonthDay[dateIdx++];

        // this logic is dependant on the BYMONTHDAYS
        // being in order (which is done by #normalizeByMonthDayRules)
        if (next >= date) {
          // if the next month day is in the future jump to it.
          lastDay = next;
        } else {
          // in this case the 'next' monthday has past
          // we must move to the month.
          nextMonth();
          continue;
        }

        // Now we can loop through the day rules to see
        // if one matches the current month date.
        for (var dayIdx = 0; dayIdx < dayLen; dayIdx++) {
          var parts = this.ruleDayOfWeek(byDay[dayIdx]);
          var pos = parts[0];
          var dow = parts[1];

          this.last.day = lastDay;
          if (this.last.isNthWeekDay(dow, pos)) {
            // when we find the valid one we can mark
            // the conditions as met and break the loop.
            // (Because we have this condition above
            //  it will also break the parent loop).
            dataIsValid = 1;
            break;
          }
        }

        // Its completely possible that the combination
        // cannot be matched in the current month.
        // When we reach the end of possible combinations
        // in the current month we iterate to the next one.
        // since dateIdx is incremented right after getting
        // "next", we don't need dateLen -1 here.
        if (!dataIsValid && dateIdx === dateLen) {
          nextMonth();
          continue;
        }
      }

      if (monthsCounter <= 0) {
        // Checked 4 years without finding a Byday that matches
        // a Bymonthday. Maybe the rule is not correct.
        throw new Error("Malformed values in BYDAY combined with BYMONTHDAY parts");
      }


      return dataIsValid;
    },

    next_month: function next_month() {
      var this_freq = (this.rule.freq == "MONTHLY");
      var data_valid = 1;

      if (this.next_hour() == 0) {
        return data_valid;
      }

      if (this.has_by_data("BYDAY") && this.has_by_data("BYMONTHDAY")) {
        data_valid = this._byDayAndMonthDay();
      } else if (this.has_by_data("BYDAY")) {
        var daysInMonth = ICAL.Time.daysInMonth(this.last.month, this.last.year);
        var setpos = 0;
        var setpos_total = 0;

        if (this.has_by_data("BYSETPOS")) {
          var last_day = this.last.day;
          for (var day = 1; day <= daysInMonth; day++) {
            this.last.day = day;
            if (this.is_day_in_byday(this.last)) {
              setpos_total++;
              if (day <= last_day) {
                setpos++;
              }
            }
          }
          this.last.day = last_day;
        }

        data_valid = 0;
        for (var day = this.last.day + 1; day <= daysInMonth; day++) {
          this.last.day = day;

          if (this.is_day_in_byday(this.last)) {
            if (!this.has_by_data("BYSETPOS") ||
                this.check_set_position(++setpos) ||
                this.check_set_position(setpos - setpos_total - 1)) {

              data_valid = 1;
              break;
            }
          }
        }

        if (day > daysInMonth) {
          this.last.day = 1;
          this.increment_month();

          if (this.is_day_in_byday(this.last)) {
            if (!this.has_by_data("BYSETPOS") || this.check_set_position(1)) {
              data_valid = 1;
            }
          } else {
            data_valid = 0;
          }
        }
      } else if (this.has_by_data("BYMONTHDAY")) {
        this.by_indices.BYMONTHDAY++;

        if (this.by_indices.BYMONTHDAY >= this.by_data.BYMONTHDAY.length) {
          this.by_indices.BYMONTHDAY = 0;
          this.increment_month();
        }

        var daysInMonth = ICAL.Time.daysInMonth(this.last.month, this.last.year);
        var day = this.by_data.BYMONTHDAY[this.by_indices.BYMONTHDAY];

        if (day < 0) {
          day = daysInMonth + day + 1;
        }

        if (day > daysInMonth) {
          this.last.day = 1;
          data_valid = this.is_day_in_byday(this.last);
        } else {
          this.last.day = day;
        }

      } else {
        this.increment_month();
        var daysInMonth = ICAL.Time.daysInMonth(this.last.month, this.last.year);
        if (this.by_data.BYMONTHDAY[0] > daysInMonth) {
          data_valid = 0;
        } else {
          this.last.day = this.by_data.BYMONTHDAY[0];
        }
      }

      return data_valid;
    },

    next_weekday_by_week: function next_weekday_by_week() {
      var end_of_data = 0;

      if (this.next_hour() == 0) {
        return end_of_data;
      }

      if (!this.has_by_data("BYDAY")) {
        return 1;
      }

      for (;;) {
        var tt = new ICAL.Time();
        this.by_indices.BYDAY++;

        if (this.by_indices.BYDAY == Object.keys(this.by_data.BYDAY).length) {
          this.by_indices.BYDAY = 0;
          end_of_data = 1;
        }

        var coded_day = this.by_data.BYDAY[this.by_indices.BYDAY];
        var parts = this.ruleDayOfWeek(coded_day);
        var dow = parts[1];

        dow -= this.rule.wkst;

        if (dow < 0) {
          dow += 7;
        }

        tt.year = this.last.year;
        tt.month = this.last.month;
        tt.day = this.last.day;

        var startOfWeek = tt.startDoyWeek(this.rule.wkst);

        if (dow + startOfWeek < 1) {
          // The selected date is in the previous year
          if (!end_of_data) {
            continue;
          }
        }

        var next = ICAL.Time.fromDayOfYear(startOfWeek + dow,
                                                  this.last.year);

        /**
         * The normalization horrors below are due to
         * the fact that when the year/month/day changes
         * it can effect the other operations that come after.
         */
        this.last.year = next.year;
        this.last.month = next.month;
        this.last.day = next.day;

        return end_of_data;
      }
    },

    next_year: function next_year() {

      if (this.next_hour() == 0) {
        return 0;
      }

      if (++this.days_index == this.days.length) {
        this.days_index = 0;
        do {
          this.increment_year(this.rule.interval);
          this.expand_year_days(this.last.year);
        } while (this.days.length == 0);
      }

      this._nextByYearDay();

      return 1;
    },

    _nextByYearDay: function _nextByYearDay() {
        var doy = this.days[this.days_index];
        var year = this.last.year;
        if (doy < 1) {
            // Time.fromDayOfYear(doy, year) indexes relative to the
            // start of the given year. That is different from the
            // semantics of BYYEARDAY where negative indexes are an
            // offset from the end of the given year.
            doy += 1;
            year += 1;
        }
        var next = ICAL.Time.fromDayOfYear(doy, year);
        this.last.day = next.day;
        this.last.month = next.month;
    },

    /**
     * @param dow (eg: '1TU', '-1MO')
     * @param {ICAL.Time.weekDay=} aWeekStart The week start weekday
     * @return [pos, numericDow] (eg: [1, 3]) numericDow is relative to aWeekStart
     */
    ruleDayOfWeek: function ruleDayOfWeek(dow, aWeekStart) {
      var matches = dow.match(/([+-]?[0-9])?(MO|TU|WE|TH|FR|SA|SU)/);
      if (matches) {
        var pos = parseInt(matches[1] || 0, 10);
        dow = ICAL.Recur.icalDayToNumericDay(matches[2], aWeekStart);
        return [pos, dow];
      } else {
        return [0, 0];
      }
    },

    next_generic: function next_generic(aRuleType, aInterval, aDateAttr,
                                        aFollowingAttr, aPreviousIncr) {
      var has_by_rule = (aRuleType in this.by_data);
      var this_freq = (this.rule.freq == aInterval);
      var end_of_data = 0;

      if (aPreviousIncr && this[aPreviousIncr]() == 0) {
        return end_of_data;
      }

      if (has_by_rule) {
        this.by_indices[aRuleType]++;
        var idx = this.by_indices[aRuleType];
        var dta = this.by_data[aRuleType];

        if (this.by_indices[aRuleType] == dta.length) {
          this.by_indices[aRuleType] = 0;
          end_of_data = 1;
        }
        this.last[aDateAttr] = dta[this.by_indices[aRuleType]];
      } else if (this_freq) {
        this["increment_" + aDateAttr](this.rule.interval);
      }

      if (has_by_rule && end_of_data && this_freq) {
        this["increment_" + aFollowingAttr](1);
      }

      return end_of_data;
    },

    increment_monthday: function increment_monthday(inc) {
      for (var i = 0; i < inc; i++) {
        var daysInMonth = ICAL.Time.daysInMonth(this.last.month, this.last.year);
        this.last.day++;

        if (this.last.day > daysInMonth) {
          this.last.day -= daysInMonth;
          this.increment_month();
        }
      }
    },

    increment_month: function increment_month() {
      this.last.day = 1;
      if (this.has_by_data("BYMONTH")) {
        this.by_indices.BYMONTH++;

        if (this.by_indices.BYMONTH == this.by_data.BYMONTH.length) {
          this.by_indices.BYMONTH = 0;
          this.increment_year(1);
        }

        this.last.month = this.by_data.BYMONTH[this.by_indices.BYMONTH];
      } else {
        if (this.rule.freq == "MONTHLY") {
          this.last.month += this.rule.interval;
        } else {
          this.last.month++;
        }

        this.last.month--;
        var years = ICAL.helpers.trunc(this.last.month / 12);
        this.last.month %= 12;
        this.last.month++;

        if (years != 0) {
          this.increment_year(years);
        }
      }
    },

    increment_year: function increment_year(inc) {
      this.last.year += inc;
    },

    increment_generic: function increment_generic(inc, aDateAttr,
                                                  aFactor, aNextIncrement) {
      this.last[aDateAttr] += inc;
      var nextunit = ICAL.helpers.trunc(this.last[aDateAttr] / aFactor);
      this.last[aDateAttr] %= aFactor;
      if (nextunit != 0) {
        this["increment_" + aNextIncrement](nextunit);
      }
    },

    has_by_data: function has_by_data(aRuleType) {
      return (aRuleType in this.rule.parts);
    },

    expand_year_days: function expand_year_days(aYear) {
      var t = new ICAL.Time();
      this.days = [];

      // We need our own copy with a few keys set
      var parts = {};
      var rules = ["BYDAY", "BYWEEKNO", "BYMONTHDAY", "BYMONTH", "BYYEARDAY"];
      for (var p in rules) {
        /* istanbul ignore else */
        if (rules.hasOwnProperty(p)) {
          var part = rules[p];
          if (part in this.rule.parts) {
            parts[part] = this.rule.parts[part];
          }
        }
      }

      if ("BYMONTH" in parts && "BYWEEKNO" in parts) {
        var valid = 1;
        var validWeeks = {};
        t.year = aYear;
        t.isDate = true;

        for (var monthIdx = 0; monthIdx < this.by_data.BYMONTH.length; monthIdx++) {
          var month = this.by_data.BYMONTH[monthIdx];
          t.month = month;
          t.day = 1;
          var first_week = t.weekNumber(this.rule.wkst);
          t.day = ICAL.Time.daysInMonth(month, aYear);
          var last_week = t.weekNumber(this.rule.wkst);
          for (monthIdx = first_week; monthIdx < last_week; monthIdx++) {
            validWeeks[monthIdx] = 1;
          }
        }

        for (var weekIdx = 0; weekIdx < this.by_data.BYWEEKNO.length && valid; weekIdx++) {
          var weekno = this.by_data.BYWEEKNO[weekIdx];
          if (weekno < 52) {
            valid &= validWeeks[weekIdx];
          } else {
            valid = 0;
          }
        }

        if (valid) {
          delete parts.BYMONTH;
        } else {
          delete parts.BYWEEKNO;
        }
      }

      var partCount = Object.keys(parts).length;

      if (partCount == 0) {
        var t1 = this.dtstart.clone();
        t1.year = this.last.year;
        this.days.push(t1.dayOfYear());
      } else if (partCount == 1 && "BYMONTH" in parts) {
        for (var monthkey in this.by_data.BYMONTH) {
          /* istanbul ignore if */
          if (!this.by_data.BYMONTH.hasOwnProperty(monthkey)) {
            continue;
          }
          var t2 = this.dtstart.clone();
          t2.year = aYear;
          t2.month = this.by_data.BYMONTH[monthkey];
          t2.isDate = true;
          this.days.push(t2.dayOfYear());
        }
      } else if (partCount == 1 && "BYMONTHDAY" in parts) {
        for (var monthdaykey in this.by_data.BYMONTHDAY) {
          /* istanbul ignore if */
          if (!this.by_data.BYMONTHDAY.hasOwnProperty(monthdaykey)) {
            continue;
          }
          var t3 = this.dtstart.clone();
          var day_ = this.by_data.BYMONTHDAY[monthdaykey];
          if (day_ < 0) {
            var daysInMonth = ICAL.Time.daysInMonth(t3.month, aYear);
            day_ = day_ + daysInMonth + 1;
          }
          t3.day = day_;
          t3.year = aYear;
          t3.isDate = true;
          this.days.push(t3.dayOfYear());
        }
      } else if (partCount == 2 &&
                 "BYMONTHDAY" in parts &&
                 "BYMONTH" in parts) {
        for (var monthkey in this.by_data.BYMONTH) {
          /* istanbul ignore if */
          if (!this.by_data.BYMONTH.hasOwnProperty(monthkey)) {
            continue;
          }
          var month_ = this.by_data.BYMONTH[monthkey];
          var daysInMonth = ICAL.Time.daysInMonth(month_, aYear);
          for (var monthdaykey in this.by_data.BYMONTHDAY) {
            /* istanbul ignore if */
            if (!this.by_data.BYMONTHDAY.hasOwnProperty(monthdaykey)) {
              continue;
            }
            var day_ = this.by_data.BYMONTHDAY[monthdaykey];
            if (day_ < 0) {
              day_ = day_ + daysInMonth + 1;
            }
            t.day = day_;
            t.month = month_;
            t.year = aYear;
            t.isDate = true;

            this.days.push(t.dayOfYear());
          }
        }
      } else if (partCount == 1 && "BYWEEKNO" in parts) {
        // TODO unimplemented in libical
      } else if (partCount == 2 &&
                 "BYWEEKNO" in parts &&
                 "BYMONTHDAY" in parts) {
        // TODO unimplemented in libical
      } else if (partCount == 1 && "BYDAY" in parts) {
        this.days = this.days.concat(this.expand_by_day(aYear));
      } else if (partCount == 2 && "BYDAY" in parts && "BYMONTH" in parts) {
        for (var monthkey in this.by_data.BYMONTH) {
          /* istanbul ignore if */
          if (!this.by_data.BYMONTH.hasOwnProperty(monthkey)) {
            continue;
          }
          var month = this.by_data.BYMONTH[monthkey];
          var daysInMonth = ICAL.Time.daysInMonth(month, aYear);

          t.year = aYear;
          t.month = this.by_data.BYMONTH[monthkey];
          t.day = 1;
          t.isDate = true;

          var first_dow = t.dayOfWeek();
          var doy_offset = t.dayOfYear() - 1;

          t.day = daysInMonth;
          var last_dow = t.dayOfWeek();

          if (this.has_by_data("BYSETPOS")) {
            var set_pos_counter = 0;
            var by_month_day = [];
            for (var day = 1; day <= daysInMonth; day++) {
              t.day = day;
              if (this.is_day_in_byday(t)) {
                by_month_day.push(day);
              }
            }

            for (var spIndex = 0; spIndex < by_month_day.length; spIndex++) {
              if (this.check_set_position(spIndex + 1) ||
                  this.check_set_position(spIndex - by_month_day.length)) {
                this.days.push(doy_offset + by_month_day[spIndex]);
              }
            }
          } else {
            for (var daycodedkey in this.by_data.BYDAY) {
              /* istanbul ignore if */
              if (!this.by_data.BYDAY.hasOwnProperty(daycodedkey)) {
                continue;
              }
              var coded_day = this.by_data.BYDAY[daycodedkey];
              var bydayParts = this.ruleDayOfWeek(coded_day);
              var pos = bydayParts[0];
              var dow = bydayParts[1];
              var month_day;

              var first_matching_day = ((dow + 7 - first_dow) % 7) + 1;
              var last_matching_day = daysInMonth - ((last_dow + 7 - dow) % 7);

              if (pos == 0) {
                for (var day = first_matching_day; day <= daysInMonth; day += 7) {
                  this.days.push(doy_offset + day);
                }
              } else if (pos > 0) {
                month_day = first_matching_day + (pos - 1) * 7;

                if (month_day <= daysInMonth) {
                  this.days.push(doy_offset + month_day);
                }
              } else {
                month_day = last_matching_day + (pos + 1) * 7;

                if (month_day > 0) {
                  this.days.push(doy_offset + month_day);
                }
              }
            }
          }
        }
        // Return dates in order of occurrence (1,2,3,...) instead
        // of by groups of weekdays (1,8,15,...,2,9,16,...).
        this.days.sort(function(a, b) { return a - b; }); // Comparator function allows to sort numbers.
      } else if (partCount == 2 && "BYDAY" in parts && "BYMONTHDAY" in parts) {
        var expandedDays = this.expand_by_day(aYear);

        for (var daykey in expandedDays) {
          /* istanbul ignore if */
          if (!expandedDays.hasOwnProperty(daykey)) {
            continue;
          }
          var day = expandedDays[daykey];
          var tt = ICAL.Time.fromDayOfYear(day, aYear);
          if (this.by_data.BYMONTHDAY.indexOf(tt.day) >= 0) {
            this.days.push(day);
          }
        }
      } else if (partCount == 3 &&
                 "BYDAY" in parts &&
                 "BYMONTHDAY" in parts &&
                 "BYMONTH" in parts) {
        var expandedDays = this.expand_by_day(aYear);

        for (var daykey in expandedDays) {
          /* istanbul ignore if */
          if (!expandedDays.hasOwnProperty(daykey)) {
            continue;
          }
          var day = expandedDays[daykey];
          var tt = ICAL.Time.fromDayOfYear(day, aYear);

          if (this.by_data.BYMONTH.indexOf(tt.month) >= 0 &&
              this.by_data.BYMONTHDAY.indexOf(tt.day) >= 0) {
            this.days.push(day);
          }
        }
      } else if (partCount == 2 && "BYDAY" in parts && "BYWEEKNO" in parts) {
        var expandedDays = this.expand_by_day(aYear);

        for (var daykey in expandedDays) {
          /* istanbul ignore if */
          if (!expandedDays.hasOwnProperty(daykey)) {
            continue;
          }
          var day = expandedDays[daykey];
          var tt = ICAL.Time.fromDayOfYear(day, aYear);
          var weekno = tt.weekNumber(this.rule.wkst);

          if (this.by_data.BYWEEKNO.indexOf(weekno)) {
            this.days.push(day);
          }
        }
      } else if (partCount == 3 &&
                 "BYDAY" in parts &&
                 "BYWEEKNO" in parts &&
                 "BYMONTHDAY" in parts) {
        // TODO unimplemted in libical
      } else if (partCount == 1 && "BYYEARDAY" in parts) {
        this.days = this.days.concat(this.by_data.BYYEARDAY);
      } else {
        this.days = [];
      }
      return 0;
    },

    expand_by_day: function expand_by_day(aYear) {

      var days_list = [];
      var tmp = this.last.clone();

      tmp.year = aYear;
      tmp.month = 1;
      tmp.day = 1;
      tmp.isDate = true;

      var start_dow = tmp.dayOfWeek();

      tmp.month = 12;
      tmp.day = 31;
      tmp.isDate = true;

      var end_dow = tmp.dayOfWeek();
      var end_year_day = tmp.dayOfYear();

      for (var daykey in this.by_data.BYDAY) {
        /* istanbul ignore if */
        if (!this.by_data.BYDAY.hasOwnProperty(daykey)) {
          continue;
        }
        var day = this.by_data.BYDAY[daykey];
        var parts = this.ruleDayOfWeek(day);
        var pos = parts[0];
        var dow = parts[1];

        if (pos == 0) {
          var tmp_start_doy = ((dow + 7 - start_dow) % 7) + 1;

          for (var doy = tmp_start_doy; doy <= end_year_day; doy += 7) {
            days_list.push(doy);
          }

        } else if (pos > 0) {
          var first;
          if (dow >= start_dow) {
            first = dow - start_dow + 1;
          } else {
            first = dow - start_dow + 8;
          }

          days_list.push(first + (pos - 1) * 7);
        } else {
          var last;
          pos = -pos;

          if (dow <= end_dow) {
            last = end_year_day - end_dow + dow;
          } else {
            last = end_year_day - end_dow + dow - 7;
          }

          days_list.push(last - (pos - 1) * 7);
        }
      }
      return days_list;
    },

    is_day_in_byday: function is_day_in_byday(tt) {
      for (var daykey in this.by_data.BYDAY) {
        /* istanbul ignore if */
        if (!this.by_data.BYDAY.hasOwnProperty(daykey)) {
          continue;
        }
        var day = this.by_data.BYDAY[daykey];
        var parts = this.ruleDayOfWeek(day);
        var pos = parts[0];
        var dow = parts[1];
        var this_dow = tt.dayOfWeek();

        if ((pos == 0 && dow == this_dow) ||
            (tt.nthWeekDay(dow, pos) == tt.day)) {
          return 1;
        }
      }

      return 0;
    },

    /**
     * Checks if given value is in BYSETPOS.
     *
     * @private
     * @param {Numeric} aPos position to check for.
     * @return {Boolean} false unless BYSETPOS rules exist
     *                   and the given value is present in rules.
     */
    check_set_position: function check_set_position(aPos) {
      if (this.has_by_data('BYSETPOS')) {
        var idx = this.by_data.BYSETPOS.indexOf(aPos);
        // negative numbers are not false-y
        return idx !== -1;
      }
      return false;
    },

    sort_byday_rules: function icalrecur_sort_byday_rules(aRules) {
      for (var i = 0; i < aRules.length; i++) {
        for (var j = 0; j < i; j++) {
          var one = this.ruleDayOfWeek(aRules[j], this.rule.wkst)[1];
          var two = this.ruleDayOfWeek(aRules[i], this.rule.wkst)[1];

          if (one > two) {
            var tmp = aRules[i];
            aRules[i] = aRules[j];
            aRules[j] = tmp;
          }
        }
      }
    },

    check_contract_restriction: function check_contract_restriction(aRuleType, v) {
      var indexMapValue = icalrecur_iterator._indexMap[aRuleType];
      var ruleMapValue = icalrecur_iterator._expandMap[this.rule.freq][indexMapValue];
      var pass = false;

      if (aRuleType in this.by_data &&
          ruleMapValue == icalrecur_iterator.CONTRACT) {

        var ruleType = this.by_data[aRuleType];

        for (var bydatakey in ruleType) {
          /* istanbul ignore else */
          if (ruleType.hasOwnProperty(bydatakey)) {
            if (ruleType[bydatakey] == v) {
              pass = true;
              break;
            }
          }
        }
      } else {
        // Not a contracting byrule or has no data, test passes
        pass = true;
      }
      return pass;
    },

    check_contracting_rules: function check_contracting_rules() {
      var dow = this.last.dayOfWeek();
      var weekNo = this.last.weekNumber(this.rule.wkst);
      var doy = this.last.dayOfYear();

      return (this.check_contract_restriction("BYSECOND", this.last.second) &&
              this.check_contract_restriction("BYMINUTE", this.last.minute) &&
              this.check_contract_restriction("BYHOUR", this.last.hour) &&
              this.check_contract_restriction("BYDAY", ICAL.Recur.numericDayToIcalDay(dow)) &&
              this.check_contract_restriction("BYWEEKNO", weekNo) &&
              this.check_contract_restriction("BYMONTHDAY", this.last.day) &&
              this.check_contract_restriction("BYMONTH", this.last.month) &&
              this.check_contract_restriction("BYYEARDAY", doy));
    },

    setup_defaults: function setup_defaults(aRuleType, req, deftime) {
      var indexMapValue = icalrecur_iterator._indexMap[aRuleType];
      var ruleMapValue = icalrecur_iterator._expandMap[this.rule.freq][indexMapValue];

      if (ruleMapValue != icalrecur_iterator.CONTRACT) {
        if (!(aRuleType in this.by_data)) {
          this.by_data[aRuleType] = [deftime];
        }
        if (this.rule.freq != req) {
          return this.by_data[aRuleType][0];
        }
      }
      return deftime;
    },

    /**
     * Convert iterator into a serialize-able object.  Will preserve current
     * iteration sequence to ensure the seamless continuation of the recurrence
     * rule.
     * @return {Object}
     */
    toJSON: function() {
      var result = Object.create(null);

      result.initialized = this.initialized;
      result.rule = this.rule.toJSON();
      result.dtstart = this.dtstart.toJSON();
      result.by_data = this.by_data;
      result.days = this.days;
      result.last = this.last.toJSON();
      result.by_indices = this.by_indices;
      result.occurrence_number = this.occurrence_number;

      return result;
    }
  };

  icalrecur_iterator._indexMap = {
    "BYSECOND": 0,
    "BYMINUTE": 1,
    "BYHOUR": 2,
    "BYDAY": 3,
    "BYMONTHDAY": 4,
    "BYYEARDAY": 5,
    "BYWEEKNO": 6,
    "BYMONTH": 7,
    "BYSETPOS": 8
  };

  icalrecur_iterator._expandMap = {
    "SECONDLY": [1, 1, 1, 1, 1, 1, 1, 1],
    "MINUTELY": [2, 1, 1, 1, 1, 1, 1, 1],
    "HOURLY": [2, 2, 1, 1, 1, 1, 1, 1],
    "DAILY": [2, 2, 2, 1, 1, 1, 1, 1],
    "WEEKLY": [2, 2, 2, 2, 3, 3, 1, 1],
    "MONTHLY": [2, 2, 2, 2, 2, 3, 3, 1],
    "YEARLY": [2, 2, 2, 2, 2, 2, 2, 2]
  };
  icalrecur_iterator.UNKNOWN = 0;
  icalrecur_iterator.CONTRACT = 1;
  icalrecur_iterator.EXPAND = 2;
  icalrecur_iterator.ILLEGAL = 3;

  return icalrecur_iterator;

}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.RecurExpansion = (function() {
  function formatTime(item) {
    return ICAL.helpers.formatClassType(item, ICAL.Time);
  }

  function compareTime(a, b) {
    return a.compare(b);
  }

  function isRecurringComponent(comp) {
    return comp.hasProperty('rdate') ||
           comp.hasProperty('rrule') ||
           comp.hasProperty('recurrence-id');
  }

  /**
   * @classdesc
   * Primary class for expanding recurring rules.  Can take multiple rrules,
   * rdates, exdate(s) and iterate (in order) over each next occurrence.
   *
   * Once initialized this class can also be serialized saved and continue
   * iteration from the last point.
   *
   * NOTE: it is intended that this class is to be used
   *       with ICAL.Event which handles recurrence exceptions.
   *
   * @example
   * // assuming event is a parsed ical component
   * var event;
   *
   * var expand = new ICAL.RecurExpansion({
   *   component: event,
   *   dtstart: event.getFirstPropertyValue('dtstart')
   * });
   *
   * // remember there are infinite rules
   * // so its a good idea to limit the scope
   * // of the iterations then resume later on.
   *
   * // next is always an ICAL.Time or null
   * var next;
   *
   * while (someCondition && (next = expand.next())) {
   *   // do something with next
   * }
   *
   * // save instance for later
   * var json = JSON.stringify(expand);
   *
   * //...
   *
   * // NOTE: if the component's properties have
   * //       changed you will need to rebuild the
   * //       class and start over. This only works
   * //       when the component's recurrence info is the same.
   * var expand = new ICAL.RecurExpansion(JSON.parse(json));
   *
   * @description
   * The options object can be filled with the specified initial values. It can
   * also contain additional members, as a result of serializing a previous
   * expansion state, as shown in the example.
   *
   * @class
   * @alias ICAL.RecurExpansion
   * @param {Object} options
   *        Recurrence expansion options
   * @param {ICAL.Time} options.dtstart
   *        Start time of the event
   * @param {ICAL.Component=} options.component
   *        Component for expansion, required if not resuming.
   */
  function RecurExpansion(options) {
    this.ruleDates = [];
    this.exDates = [];
    this.fromData(options);
  }

  RecurExpansion.prototype = {
    /**
     * True when iteration is fully completed.
     * @type {Boolean}
     */
    complete: false,

    /**
     * Array of rrule iterators.
     *
     * @type {ICAL.RecurIterator[]}
     * @private
     */
    ruleIterators: null,

    /**
     * Array of rdate instances.
     *
     * @type {ICAL.Time[]}
     * @private
     */
    ruleDates: null,

    /**
     * Array of exdate instances.
     *
     * @type {ICAL.Time[]}
     * @private
     */
    exDates: null,

    /**
     * Current position in ruleDates array.
     * @type {Number}
     * @private
     */
    ruleDateInc: 0,

    /**
     * Current position in exDates array
     * @type {Number}
     * @private
     */
    exDateInc: 0,

    /**
     * Current negative date.
     *
     * @type {ICAL.Time}
     * @private
     */
    exDate: null,

    /**
     * Current additional date.
     *
     * @type {ICAL.Time}
     * @private
     */
    ruleDate: null,

    /**
     * Start date of recurring rules.
     *
     * @type {ICAL.Time}
     */
    dtstart: null,

    /**
     * Last expanded time
     *
     * @type {ICAL.Time}
     */
    last: null,

    /**
     * Initialize the recurrence expansion from the data object. The options
     * object may also contain additional members, see the
     * {@link ICAL.RecurExpansion constructor} for more details.
     *
     * @param {Object} options
     *        Recurrence expansion options
     * @param {ICAL.Time} options.dtstart
     *        Start time of the event
     * @param {ICAL.Component=} options.component
     *        Component for expansion, required if not resuming.
     */
    fromData: function(options) {
      var start = ICAL.helpers.formatClassType(options.dtstart, ICAL.Time);

      if (!start) {
        throw new Error('.dtstart (ICAL.Time) must be given');
      } else {
        this.dtstart = start;
      }

      if (options.component) {
        this._init(options.component);
      } else {
        this.last = formatTime(options.last) || start.clone();

        if (!options.ruleIterators) {
          throw new Error('.ruleIterators or .component must be given');
        }

        this.ruleIterators = options.ruleIterators.map(function(item) {
          return ICAL.helpers.formatClassType(item, ICAL.RecurIterator);
        });

        this.ruleDateInc = options.ruleDateInc;
        this.exDateInc = options.exDateInc;

        if (options.ruleDates) {
          this.ruleDates = options.ruleDates.map(formatTime);
          this.ruleDate = this.ruleDates[this.ruleDateInc];
        }

        if (options.exDates) {
          this.exDates = options.exDates.map(formatTime);
          this.exDate = this.exDates[this.exDateInc];
        }

        if (typeof(options.complete) !== 'undefined') {
          this.complete = options.complete;
        }
      }
    },

    /**
     * Retrieve the next occurrence in the series.
     * @return {ICAL.Time}
     */
    next: function() {
      var iter;
      var ruleOfDay;
      var next;
      var compare;

      var maxTries = 500;
      var currentTry = 0;

      while (true) {
        if (currentTry++ > maxTries) {
          throw new Error(
            'max tries have occured, rule may be impossible to forfill.'
          );
        }

        next = this.ruleDate;
        iter = this._nextRecurrenceIter(this.last);

        // no more matches
        // because we increment the rule day or rule
        // _after_ we choose a value this should be
        // the only spot where we need to worry about the
        // end of events.
        if (!next && !iter) {
          // there are no more iterators or rdates
          this.complete = true;
          break;
        }

        // no next rule day or recurrence rule is first.
        if (!next || (iter && next.compare(iter.last) > 0)) {
          // must be cloned, recur will reuse the time element.
          next = iter.last.clone();
          // move to next so we can continue
          iter.next();
        }

        // if the ruleDate is still next increment it.
        if (this.ruleDate === next) {
          this._nextRuleDay();
        }

        this.last = next;

        // check the negative rules
        if (this.exDate) {
          compare = this.exDate.compare(this.last);

          if (compare < 0) {
            this._nextExDay();
          }

          // if the current rule is excluded skip it.
          if (compare === 0) {
            this._nextExDay();
            continue;
          }
        }

        //XXX: The spec states that after we resolve the final
        //     list of dates we execute exdate this seems somewhat counter
        //     intuitive to what I have seen most servers do so for now
        //     I exclude based on the original date not the one that may
        //     have been modified by the exception.
        return this.last;
      }
    },

    /**
     * Converts object into a serialize-able format. This format can be passed
     * back into the expansion to resume iteration.
     * @return {Object}
     */
    toJSON: function() {
      function toJSON(item) {
        return item.toJSON();
      }

      var result = Object.create(null);
      result.ruleIterators = this.ruleIterators.map(toJSON);

      if (this.ruleDates) {
        result.ruleDates = this.ruleDates.map(toJSON);
      }

      if (this.exDates) {
        result.exDates = this.exDates.map(toJSON);
      }

      result.ruleDateInc = this.ruleDateInc;
      result.exDateInc = this.exDateInc;
      result.last = this.last.toJSON();
      result.dtstart = this.dtstart.toJSON();
      result.complete = this.complete;

      return result;
    },

    /**
     * Extract all dates from the properties in the given component. The
     * properties will be filtered by the property name.
     *
     * @private
     * @param {ICAL.Component} component        The component to search in
     * @param {String} propertyName             The property name to search for
     * @return {ICAL.Time[]}                    The extracted dates.
     */
    _extractDates: function(component, propertyName) {
      function handleProp(prop) {
        idx = ICAL.helpers.binsearchInsert(
          result,
          prop,
          compareTime
        );

        // ordered insert
        result.splice(idx, 0, prop);
      }

      var result = [];
      var props = component.getAllProperties(propertyName);
      var len = props.length;
      var i = 0;
      var prop;

      var idx;

      for (; i < len; i++) {
        props[i].getValues().forEach(handleProp);
      }

      return result;
    },

    /**
     * Initialize the recurrence expansion.
     *
     * @private
     * @param {ICAL.Component} component    The component to initialize from.
     */
    _init: function(component) {
      this.ruleIterators = [];

      this.last = this.dtstart.clone();

      // to provide api consistency non-recurring
      // events can also use the iterator though it will
      // only return a single time.
      if (!isRecurringComponent(component)) {
        this.ruleDate = this.last.clone();
        this.complete = true;
        return;
      }

      if (component.hasProperty('rdate')) {
        this.ruleDates = this._extractDates(component, 'rdate');

        // special hack for cases where first rdate is prior
        // to the start date. We only check for the first rdate.
        // This is mostly for google's crazy recurring date logic
        // (contacts birthdays).
        if ((this.ruleDates[0]) &&
            (this.ruleDates[0].compare(this.dtstart) < 0)) {

          this.ruleDateInc = 0;
          this.last = this.ruleDates[0].clone();
        } else {
          this.ruleDateInc = ICAL.helpers.binsearchInsert(
            this.ruleDates,
            this.last,
            compareTime
          );
        }

        this.ruleDate = this.ruleDates[this.ruleDateInc];
      }

      if (component.hasProperty('rrule')) {
        var rules = component.getAllProperties('rrule');
        var i = 0;
        var len = rules.length;

        var rule;
        var iter;

        for (; i < len; i++) {
          rule = rules[i].getFirstValue();
          iter = rule.iterator(this.dtstart);
          this.ruleIterators.push(iter);

          // increment to the next occurrence so future
          // calls to next return times beyond the initial iteration.
          // XXX: I find this suspicious might be a bug?
          iter.next();
        }
      }

      if (component.hasProperty('exdate')) {
        this.exDates = this._extractDates(component, 'exdate');
        // if we have a .last day we increment the index to beyond it.
        this.exDateInc = ICAL.helpers.binsearchInsert(
          this.exDates,
          this.last,
          compareTime
        );

        this.exDate = this.exDates[this.exDateInc];
      }
    },

    /**
     * Advance to the next exdate
     * @private
     */
    _nextExDay: function() {
      this.exDate = this.exDates[++this.exDateInc];
    },

    /**
     * Advance to the next rule date
     * @private
     */
    _nextRuleDay: function() {
      this.ruleDate = this.ruleDates[++this.ruleDateInc];
    },

    /**
     * Find and return the recurrence rule with the most recent event and
     * return it.
     *
     * @private
     * @return {?ICAL.RecurIterator}    Found iterator.
     */
    _nextRecurrenceIter: function() {
      var iters = this.ruleIterators;

      if (iters.length === 0) {
        return null;
      }

      var len = iters.length;
      var iter;
      var iterTime;
      var iterIdx = 0;
      var chosenIter;

      // loop through each iterator
      for (; iterIdx < len; iterIdx++) {
        iter = iters[iterIdx];
        iterTime = iter.last;

        // if iteration is complete
        // then we must exclude it from
        // the search and remove it.
        if (iter.completed) {
          len--;
          if (iterIdx !== 0) {
            iterIdx--;
          }
          iters.splice(iterIdx, 1);
          continue;
        }

        // find the most recent possible choice
        if (!chosenIter || chosenIter.last.compare(iterTime) > 0) {
          // that iterator is saved
          chosenIter = iter;
        }
      }

      // the chosen iterator is returned but not mutated
      // this iterator contains the most recent event.
      return chosenIter;
    }
  };

  return RecurExpansion;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.Event = (function() {

  /**
   * @classdesc
   * ICAL.js is organized into multiple layers. The bottom layer is a raw jCal
   * object, followed by the component/property layer. The highest level is the
   * event representation, which this class is part of. See the
   * {@tutorial layers} guide for more details.
   *
   * @class
   * @alias ICAL.Event
   * @param {ICAL.Component=} component         The ICAL.Component to base this event on
   * @param {Object} options                    Options for this event
   * @param {Boolean} options.strictExceptions
   *          When true, will verify exceptions are related by their UUID
   * @param {Array<ICAL.Component|ICAL.Event>} options.exceptions
   *          Exceptions to this event, either as components or events. If not
   *            specified exceptions will automatically be set in relation of
   *            component's parent
   */
  function Event(component, options) {
    if (!(component instanceof ICAL.Component)) {
      options = component;
      component = null;
    }

    if (component) {
      this.component = component;
    } else {
      this.component = new ICAL.Component('vevent');
    }

    this._rangeExceptionCache = Object.create(null);
    this.exceptions = Object.create(null);
    this.except = [];
    this.rangeExceptions = [];

    if (options && options.strictExceptions) {
      this.strictExceptions = options.strictExceptions;
    }

    if (options && options.exceptions) {
      options.exceptions.forEach(this.relateException, this);
    } else if (this.component.parent && !this.isRecurrenceException()) {
      this.component.parent.getAllSubcomponents('vevent').forEach(function(event) {
        if (event.hasProperty('recurrence-id')) {
          this.relateException(event);
        }
      }, this);
    }
  }

  Event.prototype = {

    THISANDFUTURE: 'THISANDFUTURE',

    /**
     * List of related event exceptions.
     *
     * @type {ICAL.Event[]}
     */
    exceptions: null,
    except: [],

    /**
     * When true, will verify exceptions are related by their UUID.
     *
     * @type {Boolean}
     */
    strictExceptions: false,

    /**
     * Relates a given event exception to this object.  If the given component
     * does not share the UID of this event it cannot be related and will throw
     * an exception.
     *
     * If this component is an exception it cannot have other exceptions
     * related to it.
     *
     * @param {ICAL.Component|ICAL.Event} obj       Component or event
     */
    relateException: function(obj) {
      if (this.isRecurrenceException()) {
        throw new Error('cannot relate exception to exceptions');
      }

      if (obj instanceof ICAL.Component) {
        obj = new ICAL.Event(obj);
      }

      if (this.strictExceptions && obj.uid !== this.uid) {
        //throw new Error('attempted to relate unrelated exception');
      }
      else{
        this.except.push(obj);
      }

      var id = obj.recurrenceId.toString();

      // we don't sort or manage exceptions directly
      // here the recurrence expander handles that.
      this.exceptions[id] = obj;

      // index RANGE=THISANDFUTURE exceptions so we can
      // look them up later in getOccurrenceDetails.
      if (obj.modifiesFuture()) {
        var item = [
          obj.recurrenceId.toUnixTime(), id
        ];

        // we keep them sorted so we can find the nearest
        // value later on...
        var idx = ICAL.helpers.binsearchInsert(
          this.rangeExceptions,
          item,
          compareRangeException
        );

        this.rangeExceptions.splice(idx, 0, item);
      }
    },

    /**
     * Checks if this record is an exception and has the RANGE=THISANDFUTURE
     * value.
     *
     * @return {Boolean}        True, when exception is within range
     */
    modifiesFuture: function() {
      if (!this.component.hasProperty('recurrence-id')) {
        return false;
      }

      var range = this.component.getFirstProperty('recurrence-id').getParameter('range');
      return range === this.THISANDFUTURE;
    },

    /**
     * Finds the range exception nearest to the given date.
     *
     * @param {ICAL.Time} time usually an occurrence time of an event
     * @return {?ICAL.Event} the related event/exception or null
     */
    findRangeException: function(time) {
      if (!this.rangeExceptions.length) {
        return null;
      }

      var utc = time.toUnixTime();
      var idx = ICAL.helpers.binsearchInsert(
        this.rangeExceptions,
        [utc],
        compareRangeException
      );

      idx -= 1;

      // occurs before
      if (idx < 0) {
        return null;
      }

      var rangeItem = this.rangeExceptions[idx];

      /* istanbul ignore next: sanity check only */
      if (utc < rangeItem[0]) {
        return null;
      }

      return rangeItem[1];
    },

    /**
     * This object is returned by {@link ICAL.Event#getOccurrenceDetails getOccurrenceDetails}
     *
     * @typedef {Object} occurrenceDetails
     * @memberof ICAL.Event
     * @property {ICAL.Time} recurrenceId       The passed in recurrence id
     * @property {ICAL.Event} item              The occurrence
     * @property {ICAL.Time} startDate          The start of the occurrence
     * @property {ICAL.Time} endDate            The end of the occurrence
     */

    /**
     * Returns the occurrence details based on its start time.  If the
     * occurrence has an exception will return the details for that exception.
     *
     * NOTE: this method is intend to be used in conjunction
     *       with the {@link ICAL.Event#iterator iterator} method.
     *
     * @param {ICAL.Time} occurrence time occurrence
     * @return {ICAL.Event.occurrenceDetails} Information about the occurrence
     */
    getOccurrenceDetails: function(occurrence) {
      var id = occurrence.toString();
      var utcId = occurrence.convertToZone(ICAL.Timezone.utcTimezone).toString();
      var item;
      var result = {
        //XXX: Clone?
        recurrenceId: occurrence
      };

      if (id in this.exceptions) {
        item = result.item = this.exceptions[id];
        result.startDate = item.startDate;
        result.endDate = item.endDate;
        result.item = item;
      } else if (utcId in this.exceptions) {
        item = this.exceptions[utcId];
        result.startDate = item.startDate;
        result.endDate = item.endDate;
        result.item = item;
      } else {
        // range exceptions (RANGE=THISANDFUTURE) have a
        // lower priority then direct exceptions but
        // must be accounted for first. Their item is
        // always the first exception with the range prop.
        var rangeExceptionId = this.findRangeException(
          occurrence
        );
        var end;

        if (rangeExceptionId) {
          var exception = this.exceptions[rangeExceptionId];

          // range exception must modify standard time
          // by the difference (if any) in start/end times.
          result.item = exception;

          var startDiff = this._rangeExceptionCache[rangeExceptionId];

          if (!startDiff) {
            var original = exception.recurrenceId.clone();
            var newStart = exception.startDate.clone();

            // zones must be same otherwise subtract may be incorrect.
            original.zone = newStart.zone;
            startDiff = newStart.subtractDate(original);

            this._rangeExceptionCache[rangeExceptionId] = startDiff;
          }

          var start = occurrence.clone();
          start.zone = exception.startDate.zone;
          start.addDuration(startDiff);

          end = start.clone();
          end.addDuration(exception.duration);

          result.startDate = start;
          result.endDate = end;
        } else {
          // no range exception standard expansion
          end = occurrence.clone();
          end.addDuration(this.duration);

          result.endDate = end;
          result.startDate = occurrence;
          result.item = this;
        }
      }

      return result;
    },

    /**
     * Builds a recur expansion instance for a specific point in time (defaults
     * to startDate).
     *
     * @param {ICAL.Time} startTime     Starting point for expansion
     * @return {ICAL.RecurExpansion}    Expansion object
     */
    iterator: function(startTime) {
      return new ICAL.RecurExpansion({
        component: this.component,
        dtstart: startTime || this.startDate
      });
    },

    /**
     * Checks if the event is recurring
     *
     * @return {Boolean}        True, if event is recurring
     */
    isRecurring: function() {
      var comp = this.component;
      return comp.hasProperty('rrule') || comp.hasProperty('rdate');
    },

    /**
     * Checks if the event describes a recurrence exception. See
     * {@tutorial terminology} for details.
     *
     * @return {Boolean}    True, if the even describes a recurrence exception
     */
    isRecurrenceException: function() {
      return this.component.hasProperty('recurrence-id');
    },

    /**
     * Returns the types of recurrences this event may have.
     *
     * Returned as an object with the following possible keys:
     *
     *    - YEARLY
     *    - MONTHLY
     *    - WEEKLY
     *    - DAILY
     *    - MINUTELY
     *    - SECONDLY
     *
     * @return {Object.<ICAL.Recur.frequencyValues, Boolean>}
     *          Object of recurrence flags
     */
    getRecurrenceTypes: function() {
      var rules = this.component.getAllProperties('rrule');
      var i = 0;
      var len = rules.length;
      var result = Object.create(null);

      for (; i < len; i++) {
        var value = rules[i].getFirstValue();
        result[value.freq] = true;
      }

      return result;
    },

    /**
     * The uid of this event
     * @type {String}
     */
    get uid() {
      return this._firstProp('uid');
    },

    set uid(value) {
      this._setProp('uid', value);
    },

    /**
     * The start date
     * @type {ICAL.Time}
     */
    get startDate() {
      return this._firstProp('dtstart');
    },

    set startDate(value) {
      this._setTime('dtstart', value);
    },

    /**
     * The end date. This can be the result directly from the property, or the
     * end date calculated from start date and duration. Setting the property
     * will remove any duration properties.
     * @type {ICAL.Time}
     */
    get endDate() {
      var endDate = this._firstProp('dtend');
      if (!endDate) {
          var duration = this._firstProp('duration');
          endDate = this.startDate.clone();
          if (duration) {
              endDate.addDuration(duration);
          } else if (endDate.isDate) {
              endDate.day += 1;
          }
      }
      return endDate;
    },

    set endDate(value) {
      if (this.component.hasProperty('duration')) {
        this.component.removeProperty('duration');
      }
      this._setTime('dtend', value);
    },

    /**
     * The duration. This can be the result directly from the property, or the
     * duration calculated from start date and end date. Setting the property
     * will remove any `dtend` properties.
     * @type {ICAL.Duration}
     */
    get duration() {
      var duration = this._firstProp('duration');
      if (!duration) {
        return this.endDate.subtractDateTz(this.startDate);
      }
      return duration;
    },

    set duration(value) {
      if (this.component.hasProperty('dtend')) {
        this.component.removeProperty('dtend');
      }

      this._setProp('duration', value);
    },

    /**
     * The location of the event.
     * @type {String}
     */
    get location() {
      return this._firstProp('location');
    },

    set location(value) {
      return this._setProp('location', value);
    },

    /**
     * The attendees in the event
     * @type {ICAL.Property[]}
     * @readonly
     */
    get attendees() {
      //XXX: This is way lame we should have a better
      //     data structure for this later.
      return this.component.getAllProperties('attendee');
    },


    /**
     * The event summary
     * @type {String}
     */
    get summary() {
      return this._firstProp('summary');
    },

    set summary(value) {
      this._setProp('summary', value);
    },

    /**
     * The event description.
     * @type {String}
     */
    get description() {
      return this._firstProp('description');
    },

    set description(value) {
      this._setProp('description', value);
    },

    /**
     * The organizer value as an uri. In most cases this is a mailto: uri, but
     * it can also be something else, like urn:uuid:...
     * @type {String}
     */
    get organizer() {
      return this._firstProp('organizer');
    },

    set organizer(value) {
      this._setProp('organizer', value);
    },

    /**
     * The sequence value for this event. Used for scheduling
     * see {@tutorial terminology}.
     * @type {Number}
     */
    get sequence() {
      return this._firstProp('sequence');
    },

    set sequence(value) {
      this._setProp('sequence', value);
    },

    /**
     * The recurrence id for this event. See {@tutorial terminology} for details.
     * @type {ICAL.Time}
     */
    get recurrenceId() {
      return this._firstProp('recurrence-id');
    },

    set recurrenceId(value) {
      this._setTime('recurrence-id', value);
    },

    /**
     * Set/update a time property's value.
     * This will also update the TZID of the property.
     *
     * TODO: this method handles the case where we are switching
     * from a known timezone to an implied timezone (one without TZID).
     * This does _not_ handle the case of moving between a known
     *  (by TimezoneService) timezone to an unknown timezone...
     *
     * We will not add/remove/update the VTIMEZONE subcomponents
     *  leading to invalid ICAL data...
     * @private
     * @param {String} propName     The property name
     * @param {ICAL.Time} time      The time to set
     */
    _setTime: function(propName, time) {
      var prop = this.component.getFirstProperty(propName);

      if (!prop) {
        prop = new ICAL.Property(propName);
        this.component.addProperty(prop);
      }

      // utc and local don't get a tzid
      if (
        time.zone === ICAL.Timezone.localTimezone ||
        time.zone === ICAL.Timezone.utcTimezone
      ) {
        // remove the tzid
        prop.removeParameter('tzid');
      } else {
        prop.setParameter('tzid', time.zone.tzid);
      }

      prop.setValue(time);
    },

    _setProp: function(name, value) {
      this.component.updatePropertyWithValue(name, value);
    },

    _firstProp: function(name) {
      return this.component.getFirstPropertyValue(name);
    },

    /**
     * The string representation of this event.
     * @return {String}
     */
    toString: function() {
      return this.component.toString();
    }

  };

  function compareRangeException(a, b) {
    if (a[0] > b[0]) return 1;
    if (b[0] > a[0]) return -1;
    return 0;
  }

  return Event;
}());
/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 * Portions Copyright (C) Philipp Kewisch, 2011-2015 */


/**
 * This symbol is further described later on
 * @ignore
 */
ICAL.ComponentParser = (function() {
  /**
   * @classdesc
   * The ComponentParser is used to process a String or jCal Object,
   * firing callbacks for various found components, as well as completion.
   *
   * @example
   * var options = {
   *   // when false no events will be emitted for type
   *   parseEvent: true,
   *   parseTimezone: true
   * };
   *
   * var parser = new ICAL.ComponentParser(options);
   *
   * parser.onevent(eventComponent) {
   *   //...
   * }
   *
   * // ontimezone, etc...
   *
   * parser.oncomplete = function() {
   *
   * };
   *
   * parser.process(stringOrComponent);
   *
   * @class
   * @alias ICAL.ComponentParser
   * @param {Object=} options        Component parser options
   * @param {Boolean} options.parseEvent        Whether events should be parsed
   * @param {Boolean} options.parseTimezeone    Whether timezones should be parsed
   */
  function ComponentParser(options) {
    if (typeof(options) === 'undefined') {
      options = {};
    }

    var key;
    for (key in options) {
      /* istanbul ignore else */
      if (options.hasOwnProperty(key)) {
        this[key] = options[key];
      }
    }
  }

  ComponentParser.prototype = {

    /**
     * When true, parse events
     *
     * @type {Boolean}
     */
    parseEvent: true,

    /**
     * When true, parse timezones
     *
     * @type {Boolean}
     */
    parseTimezone: true,


    /* SAX like events here for reference */

    /**
     * Fired when parsing is complete
     * @callback
     */
    oncomplete: /* istanbul ignore next */ function() {},

    /**
     * Fired if an error occurs during parsing.
     *
     * @callback
     * @param {Error} err details of error
     */
    onerror: /* istanbul ignore next */ function(err) {},

    /**
     * Fired when a top level component (VTIMEZONE) is found
     *
     * @callback
     * @param {ICAL.Timezone} component     Timezone object
     */
    ontimezone: /* istanbul ignore next */ function(component) {},

    /**
     * Fired when a top level component (VEVENT) is found.
     *
     * @callback
     * @param {ICAL.Event} component    Top level component
     */
    onevent: /* istanbul ignore next */ function(component) {},

    /**
     * Process a string or parse ical object.  This function itself will return
     * nothing but will start the parsing process.
     *
     * Events must be registered prior to calling this method.
     *
     * @param {ICAL.Component|String|Object} ical      The component to process,
     *        either in its final form, as a jCal Object, or string representation
     */
    process: function(ical) {
      //TODO: this is sync now in the future we will have a incremental parser.
      if (typeof(ical) === 'string') {
        ical = ICAL.parse(ical);
      }

      if (!(ical instanceof ICAL.Component)) {
        ical = new ICAL.Component(ical);
      }

      var components = ical.getAllSubcomponents();
      var i = 0;
      var len = components.length;
      var component;

      for (; i < len; i++) {
        component = components[i];

        switch (component.name) {
          case 'vtimezone':
            if (this.parseTimezone) {
              var tzid = component.getFirstPropertyValue('tzid');
              if (tzid) {
                this.ontimezone(new ICAL.Timezone({
                  tzid: tzid,
                  component: component
                }));
              }
            }
            break;
          case 'vevent':
            if (this.parseEvent) {
              this.onevent(new ICAL.Event(component));
            }
            break;
          default:
            continue;
        }
      }

      //XXX: ideally we should do a "nextTick" here
      //     so in all cases this is actually async.
      this.oncomplete();
    }
  };

  return ComponentParser;
}());

// Source: vendor/tzid.js
var tzids = [
"Africa/Abidjan",
"Africa/Accra",
"Africa/Addis_Ababa",
"Africa/Algiers",
"Africa/Asmara",
"Africa/Bamako",
"Africa/Bangui",
"Africa/Banjul",
"Africa/Bissau",
"Africa/Blantyre",
"Africa/Brazzaville",
"Africa/Bujumbura",
"Africa/Cairo",
"Africa/Casablanca",
"Africa/Ceuta",
"Africa/Conakry",
"Africa/Dakar",
"Africa/Dar_es_Salaam",
"Africa/Djibouti",
"Africa/Douala",
"Africa/El_Aaiun",
"Africa/Freetown",
"Africa/Gaborone",
"Africa/Harare",
"Africa/Johannesburg",
"Africa/Juba",
"Africa/Kampala",
"Africa/Khartoum",
"Africa/Kigali",
"Africa/Kinshasa",
"Africa/Lagos",
"Africa/Libreville",
"Africa/Lome",
"Africa/Luanda",
"Africa/Lubumbashi",
"Africa/Lusaka",
"Africa/Malabo",
"Africa/Maputo",
"Africa/Maseru",
"Africa/Mbabane",
"Africa/Mogadishu",
"Africa/Monrovia",
"Africa/Nairobi",
"Africa/Ndjamena",
"Africa/Niamey",
"Africa/Nouakchott",
"Africa/Ouagadougou",
"Africa/Porto-Novo",
"Africa/Sao_Tome",
"Africa/Timbuktu",
"Africa/Tripoli",
"Africa/Tunis",
"Africa/Windhoek",
"America/Adak",
"America/Anchorage",
"America/Anguilla",
"America/Antigua",
"America/Araguaina",
"America/Argentina/Buenos_Aires",
"America/Argentina/Catamarca",
"America/Argentina/ComodRivadavia",
"America/Argentina/Cordoba",
"America/Argentina/Jujuy",
"America/Argentina/La_Rioja",
"America/Argentina/Mendoza",
"America/Argentina/Rio_Gallegos",
"America/Argentina/Salta",
"America/Argentina/San_Juan",
"America/Argentina/San_Luis",
"America/Argentina/Tucuman",
"America/Argentina/Ushuaia",
"America/Aruba",
"America/Asuncion",
"America/Atikokan",
"America/Atka",
"America/Bahia",
"America/Bahia_Banderas",
"America/Barbados",
"America/Belem",
"America/Belize",
"America/Blanc-Sablon",
"America/Boa_Vista",
"America/Bogota",
"America/Boise",
"America/Buenos_Aires",
"America/Cambridge_Bay",
"America/Campo_Grande",
"America/Cancun",
"America/Caracas",
"America/Catamarca",
"America/Cayenne",
"America/Cayman",
"America/Chicago",
"America/Chihuahua",
"America/Coral_Harbour",
"America/Cordoba",
"America/Costa_Rica",
"America/Creston",
"America/Cuiaba",
"America/Curacao",
"America/Danmarkshavn",
"America/Dawson",
"America/Dawson_Creek",
"America/Denver",
"America/Detroit",
"America/Dominica",
"America/Edmonton",
"America/Eirunepe",
"America/El_Salvador",
"America/Ensenada",
"America/Fort_Nelson",
"America/Fort_Wayne",
"America/Fortaleza",
"America/Glace_Bay",
"America/Godthab",
"America/Goose_Bay",
"America/Grand_Turk",
"America/Grenada",
"America/Guadeloupe",
"America/Guatemala",
"America/Guayaquil",
"America/Guyana",
"America/Halifax",
"America/Havana",
"America/Hermosillo",
"America/Indiana/Indianapolis",
"America/Indiana/Knox",
"America/Indiana/Marengo",
"America/Indiana/Petersburg",
"America/Indiana/Tell_City",
"America/Indiana/Vevay",
"America/Indiana/Vincennes",
"America/Indiana/Winamac",
"America/Indianapolis",
"America/Inuvik",
"America/Iqaluit",
"America/Jamaica",
"America/Jujuy",
"America/Juneau",
"America/Kentucky/Louisville",
"America/Kentucky/Monticello",
"America/Knox_IN",
"America/Kralendijk",
"America/La_Paz",
"America/Lima",
"America/Los_Angeles",
"America/Louisville",
"America/Lower_Princes",
"America/Maceio",
"America/Managua",
"America/Manaus",
"America/Marigot",
"America/Martinique",
"America/Matamoros",
"America/Mazatlan",
"America/Mendoza",
"America/Menominee",
"America/Merida",
"America/Metlakatla",
"America/Mexico_City",
"America/Miquelon",
"America/Moncton",
"America/Monterrey",
"America/Montevideo",
"America/Montreal",
"America/Montserrat",
"America/Nassau",
"America/New_York",
"America/Nipigon",
"America/Nome",
"America/Noronha",
"America/North_Dakota/Beulah",
"America/North_Dakota/Center",
"America/North_Dakota/New_Salem",
"America/Ojinaga",
"America/Panama",
"America/Pangnirtung",
"America/Paramaribo",
"America/Phoenix",
"America/Port_of_Spain",
"America/Port-au-Prince",
"America/Porto_Acre",
"America/Porto_Velho",
"America/Puerto_Rico",
"America/Punta_Arenas",
"America/Rainy_River",
"America/Rankin_Inlet",
"America/Recife",
"America/Regina",
"America/Resolute",
"America/Rio_Branco",
"America/Rosario",
"America/Santa_Isabel",
"America/Santarem",
"America/Santiago",
"America/Santo_Domingo",
"America/Sao_Paulo",
"America/Scoresbysund",
"America/Shiprock",
"America/Sitka",
"America/St_Barthelemy",
"America/St_Johns",
"America/St_Kitts",
"America/St_Lucia",
"America/St_Thomas",
"America/St_Vincent",
"America/Swift_Current",
"America/Tegucigalpa",
"America/Thule",
"America/Thunder_Bay",
"America/Tijuana",
"America/Toronto",
"America/Tortola",
"America/Vancouver",
"America/Virgin",
"America/Whitehorse",
"America/Winnipeg",
"America/Yakutat",
"America/Yellowknife",
"Antarctica/Casey",
"Antarctica/Davis",
"Antarctica/DumontDUrville",
"Antarctica/Macquarie",
"Antarctica/Mawson",
"Antarctica/McMurdo",
"Antarctica/Palmer",
"Antarctica/Rothera",
"Antarctica/South_Pole",
"Antarctica/Syowa",
"Antarctica/Troll",
"Antarctica/Vostok",
"Arctic/Longyearbyen",
"Asia/Aden",
"Asia/Almaty",
"Asia/Amman",
"Asia/Anadyr",
"Asia/Aqtau",
"Asia/Aqtobe",
"Asia/Ashgabat",
"Asia/Ashkhabad",
"Asia/Atyrau",
"Asia/Baghdad",
"Asia/Bahrain",
"Asia/Baku",
"Asia/Bangkok",
"Asia/Barnaul",
"Asia/Beirut",
"Asia/Bishkek",
"Asia/Brunei",
"Asia/Calcutta",
"Asia/Chita",
"Asia/Choibalsan",
"Asia/Chongqing",
"Asia/Chungking",
"Asia/Colombo",
"Asia/Dacca",
"Asia/Damascus",
"Asia/Dhaka",
"Asia/Dili",
"Asia/Dubai",
"Asia/Dushanbe",
"Asia/Famagusta",
"Asia/Gaza",
"Asia/Harbin",
"Asia/Hebron",
"Asia/Ho_Chi_Minh",
"Asia/Hong_Kong",
"Asia/Hovd",
"Asia/Irkutsk",
"Asia/Istanbul",
"Asia/Jakarta",
"Asia/Jayapura",
"Asia/Jerusalem",
"Asia/Kabul",
"Asia/Kamchatka",
"Asia/Karachi",
"Asia/Kashgar",
"Asia/Kathmandu",
"Asia/Katmandu",
"Asia/Khandyga",
"Asia/Kolkata",
"Asia/Krasnoyarsk",
"Asia/Kuala_Lumpur",
"Asia/Kuching",
"Asia/Kuwait",
"Asia/Macao",
"Asia/Macau",
"Asia/Magadan",
"Asia/Makassar",
"Asia/Manila",
"Asia/Muscat",
"Asia/Novokuznetsk",
"Asia/Novosibirsk",
"Asia/Omsk",
"Asia/Oral",
"Asia/Phnom_Penh",
"Asia/Pontianak",
"Asia/Pyongyang",
"Asia/Qatar",
"Asia/Qyzylorda",
"Asia/Rangoon",
"Asia/Riyadh",
"Asia/Saigon",
"Asia/Sakhalin",
"Asia/Samarkand",
"Asia/Seoul",
"Asia/Shanghai",
"Asia/Singapore",
"Asia/Srednekolymsk",
"Asia/Taipei",
"Asia/Tashkent",
"Asia/Tbilisi",
"Asia/Tehran",
"Asia/Tel_Aviv",
"Asia/Thimbu",
"Asia/Thimphu",
"Asia/Tokyo",
"Asia/Tomsk",
"Asia/Ujung_Pandang",
"Asia/Ulaanbaatar",
"Asia/Ulan_Bator",
"Asia/Urumqi",
"Asia/Ust-Nera",
"Asia/Vientiane",
"Asia/Vladivostok",
"Asia/Yakutsk",
"Asia/Yangon",
"Asia/Yekaterinburg",
"Asia/Yerevan",
"Atlantic/Azores",
"Atlantic/Bermuda",
"Atlantic/Canary",
"Atlantic/Cape_Verde",
"Atlantic/Faeroe",
"Atlantic/Faroe",
"Atlantic/Jan_Mayen",
"Atlantic/Madeira",
"Atlantic/Reykjavik",
"Atlantic/South_Georgia",
"Atlantic/St_Helena",
"Atlantic/Stanley",
"Australia/ACT",
"Australia/Adelaide",
"Australia/Brisbane",
"Australia/Broken_Hill",
"Australia/Canberra",
"Australia/Currie",
"Australia/Darwin",
"Australia/Eucla",
"Australia/Hobart",
"Australia/LHI",
"Australia/Lindeman",
"Australia/Lord_Howe",
"Australia/Melbourne",
"Australia/North",
"Australia/NSW",
"Australia/Perth",
"Australia/Queensland",
"Australia/South",
"Australia/Sydney",
"Australia/Tasmania",
"Australia/Victoria",
"Australia/West",
"Australia/Yancowinna",
"Brazil/Acre",
"Brazil/DeNoronha",
"Brazil/East",
"Brazil/West",
"Canada/Atlantic",
"Canada/Central",
"Canada/Eastern",
"Canada/Mountain",
"Canada/Newfoundland",
"Canada/Pacific",
"Canada/Saskatchewan",
"Canada/Yukon",
"CET",
"Chile/Continental",
"Chile/EasterIsland",
"CST6CDT",
"Cuba",
"EET",
"Egypt",
"Eire",
"EST",
"EST5EDT",
"Etc/GMT",
"Etc/GMT+0",
"Etc/GMT+1",
"Etc/GMT+10",
"Etc/GMT+11",
"Etc/GMT+12",
"Etc/GMT+2",
"Etc/GMT+3",
"Etc/GMT+4",
"Etc/GMT+5",
"Etc/GMT+6",
"Etc/GMT+7",
"Etc/GMT+8",
"Etc/GMT+9",
"Etc/GMT0",
"Etc/GMT-0",
"Etc/GMT-1",
"Etc/GMT-10",
"Etc/GMT-11",
"Etc/GMT-12",
"Etc/GMT-13",
"Etc/GMT-14",
"Etc/GMT-2",
"Etc/GMT-3",
"Etc/GMT-4",
"Etc/GMT-5",
"Etc/GMT-6",
"Etc/GMT-7",
"Etc/GMT-8",
"Etc/GMT-9",
"Etc/Greenwich",
"Etc/UCT",
"Etc/Universal",
"Etc/UTC",
"Etc/Zulu",
"Europe/Amsterdam",
"Europe/Andorra",
"Europe/Astrakhan",
"Europe/Athens",
"Europe/Belfast",
"Europe/Belgrade",
"Europe/Berlin",
"Europe/Bratislava",
"Europe/Brussels",
"Europe/Bucharest",
"Europe/Budapest",
"Europe/Busingen",
"Europe/Chisinau",
"Europe/Copenhagen",
"Europe/Dublin",
"Europe/Gibraltar",
"Europe/Guernsey",
"Europe/Helsinki",
"Europe/Isle_of_Man",
"Europe/Istanbul",
"Europe/Jersey",
"Europe/Kaliningrad",
"Europe/Kiev",
"Europe/Kirov",
"Europe/Lisbon",
"Europe/Ljubljana",
"Europe/London",
"Europe/Luxembourg",
"Europe/Madrid",
"Europe/Malta",
"Europe/Mariehamn",
"Europe/Minsk",
"Europe/Monaco",
"Europe/Moscow",
"Asia/Nicosia",
"Europe/Oslo",
"Europe/Paris",
"Europe/Podgorica",
"Europe/Prague",
"Europe/Riga",
"Europe/Rome",
"Europe/Samara",
"Europe/San_Marino",
"Europe/Sarajevo",
"Europe/Saratov",
"Europe/Simferopol",
"Europe/Skopje",
"Europe/Sofia",
"Europe/Stockholm",
"Europe/Tallinn",
"Europe/Tirane",
"Europe/Tiraspol",
"Europe/Ulyanovsk",
"Europe/Uzhgorod",
"Europe/Vaduz",
"Europe/Vatican",
"Europe/Vienna",
"Europe/Vilnius",
"Europe/Volgograd",
"Europe/Warsaw",
"Europe/Zagreb",
"Europe/Zaporozhye",
"Europe/Zurich",
"GB",
"GB-Eire",
"GMT",
"GMT+0",
"GMT0",
"GMT−0",
"Hongkong",
"HST",
"Iceland",
"Indian/Antananarivo",
"Indian/Chagos",
"Indian/Christmas",
"Indian/Cocos",
"Indian/Comoro",
"Indian/Kerguelen",
"Indian/Mahe",
"Indian/Maldives",
"Indian/Mauritius",
"Indian/Mayotte",
"Indian/Reunion",
"Iran",
"Israel",
"Jamaica",
"Japan",
"Kwajalein",
"Libya",
"MET",
"Mexico/BajaNorte",
"Mexico/BajaSur",
"Mexico/General",
"MST",
"MST7MDT",
"Navajo",
"NZ",
"NZ-CHAT",
"Pacific/Apia",
"Pacific/Auckland",
"Pacific/Bougainville",
"Pacific/Chatham",
"Pacific/Chuuk",
"Pacific/Easter",
"Pacific/Efate",
"Pacific/Enderbury",
"Pacific/Fakaofo",
"Pacific/Fiji",
"Pacific/Funafuti",
"Pacific/Galapagos",
"Pacific/Gambier",
"Pacific/Guadalcanal",
"Pacific/Guam",
"Pacific/Honolulu",
"Pacific/Johnston",
"Pacific/Kiritimati",
"Pacific/Kosrae",
"Pacific/Kwajalein",
"Pacific/Majuro",
"Pacific/Marquesas",
"Pacific/Midway",
"Pacific/Nauru",
"Pacific/Niue",
"Pacific/Norfolk",
"Pacific/Noumea",
"Pacific/Pago_Pago",
"Pacific/Palau",
"Pacific/Pitcairn",
"Pacific/Pohnpei",
"Pacific/Ponape",
"Pacific/Port_Moresby",
"Pacific/Rarotonga",
"Pacific/Saipan",
"Pacific/Samoa",
"Pacific/Tahiti",
"Pacific/Tarawa",
"Pacific/Tongatapu",
"Pacific/Truk",
"Pacific/Wake",
"Pacific/Wallis",
"Pacific/Yap",
"Poland",
"Portugal",
"PRC",
"PST8PDT",
"ROC",
"ROK",
"Singapore",
"Turkey",
"UCT",
"Universal",
"US/Alaska",
"US/Aleutian",
"US/Arizona",
"US/Central",
"US/Eastern",
"US/East-Indiana",
"US/Hawaii",
"US/Indiana-Starke",
"US/Michigan",
"US/Mountain",
"US/Pacific",
"US/Pacific-New",
"US/Samoa",
"UTC",
"WET",
"W-SU",
"Zulu"
];
//Windows Timezone names to IANA, according https://github.com/unicode-org/cldr/blob/master/common/supplemental/windowsZones.xml
tzidreplace = {
"Dateline Standard Time": "Etc/GMT+12",
"UTC-11": "Etc/GMT+11",
"Aleutian Standard Time": "America/Adak",
"Hawaiian Standard Time": "Pacific/Honolulu",
"Marquesas Standard Time": "Pacific/Marquesas",
"Alaskan Standard Time": "America/Anchorage",
"UTC-09": "Pacific/Gambier",
"Pacific Standard Time (Mexico)": "America/Tijuana",
"UTC-08": "Pacific/Pitcairn",
"Pacific Standard Time": "America/Los_Angeles",
"US Mountain Standard Time": "America/Phoenix",
"Mountain Standard Time (Mexico)": "America/Chihuahua",
"Mountain Standard Time": "America/Denver",
"Central America Standard Time": "America/Guatemala",
"Central Standard Time": "America/Chicago",
"Easter Island Standard Time": "Pacific/Easter",
"Central Standard Time (Mexico)": "America/Mexico_City",
"Canada Central Standard Time": "America/Regina",
"SA Pacific Standard Time": "America/Bogota",
"Eastern Standard Time (Mexico)": "America/Cancun",
"Eastern Standard Time": "America/New_York",
"Haiti Standard Time": "America/Port-au-Prince",
"Cuba Standard Time": "America/Havana",
"US Eastern Standard Time": "America/Indianapolis",
"Paraguay Standard Time": "America/Asuncion",
"Atlantic Standard Time": "America/Halifax",
"Venezuela Standard Time": "America/Caracas",
"Central Brazilian Standard Time": "America/Cuiaba",
"SA Western Standard Time": "America/La_Paz",
"Pacific SA Standard Time": "America/Santiago",
"Turks And Caicos Standard Time": "America/Grand_Turk",
"Newfoundland Standard Time": "America/St_Johns",
"Tocantins Standard Time": "America/Araguaina",
"E. South America Standard Time": "America/Sao_Paulo",
"SA Eastern Standard Time": "America/Cayenne",
"Argentina Standard Time": "America/Buenos_Aires",
"Greenland Standard Time": "America/Godthab",
"Montevideo Standard Time": "America/Montevideo",
"Magallanes Standard Time": "America/Punta_Arenas",
"Saint Pierre Standard Time": "America/Miquelon",
"Bahia Standard Time": "America/Bahia",
"UTC-02": "America/Noronha",
"Azores Standard Time": "Atlantic/Azores",
"Cape Verde Standard Time": "Atlantic/Cape_Verde",
"UTC": "Etc/GMT",
"GMT Standard Time": "Europe/London",
"Greenwich Standard Time": "Atlantic/Reykjavik",
"W. Europe Standard Time": "Europe/Berlin",
"Central Europe Standard Time": "Europe/Budapest",
"Romance Standard Time": "Europe/Paris",
"Morocco Standard Time": "Africa/Casablanca",
"Sao Tome Standard Time": "Africa/Sao_Tome",
"Central European Standard Time": "Europe/Warsaw",
"W. Central Africa Standard Time": "Africa/Lagos",
"Jordan Standard Time": "Asia/Amman",
"GTB Standard Time": "Europe/Bucharest",
"Middle East Standard Time": "Asia/Beirut",
"Egypt Standard Time": "Africa/Cairo",
"E. Europe Standard Time": "Europe/Chisinau",
"Syria Standard Time": "Asia/Damascus",
"West Bank Standard Time": "Asia/Hebron",
"South Africa Standard Time": "Africa/Johannesburg",
"FLE Standard Time": "Europe/Kiev",
"Israel Standard Time": "Asia/Jerusalem",
"Kaliningrad Standard Time": "Europe/Kaliningrad",
"Sudan Standard Time": "Africa/Khartoum",
"Libya Standard Time": "Africa/Tripoli",
"Namibia Standard Time": "Africa/Windhoek",
"Arabic Standard Time": "Asia/Baghdad",
"Turkey Standard Time": "Europe/Istanbul",
"Arab Standard Time": "Asia/Riyadh",
"Belarus Standard Time": "Europe/Minsk",
"Russian Standard Time": "Europe/Moscow",
"E. Africa Standard Time": "Africa/Nairobi",
"Iran Standard Time": "Asia/Tehran",
"Arabian Standard Time": "Asia/Dubai",
"Astrakhan Standard Time": "Europe/Astrakhan",
"Azerbaijan Standard Time": "Asia/Baku",
"Russia Time Zone 3": "Europe/Samara",
"Mauritius Standard Time": "Indian/Mauritius",
"Saratov Standard Time": "Europe/Saratov",
"Georgian Standard Time": "Asia/Tbilisi",
"Caucasus Standard Time": "Asia/Yerevan",
"Afghanistan Standard Time": "Asia/Kabul",
"West Asia Standard Time": "Asia/Tashkent",
"Ekaterinburg Standard Time": "Asia/Yekaterinburg",
"Pakistan Standard Time": "Asia/Karachi",
"India Standard Time": "Asia/Calcutta",
"Sri Lanka Standard Time": "Asia/Colombo",
"Nepal Standard Time": "Asia/Katmandu",
"Central Asia Standard Time": "Asia/Almaty",
"Bangladesh Standard Time": "Asia/Dhaka",
"Omsk Standard Time": "Asia/Omsk",
"Myanmar Standard Time": "Asia/Rangoon",
"SE Asia Standard Time": "Asia/Bangkok",
"Altai Standard Time": "Asia/Barnaul",
"W. Mongolia Standard Time": "Asia/Hovd",
"North Asia Standard Time": "Asia/Krasnoyarsk",
"N. Central Asia Standard Time": "Asia/Novosibirsk",
"Tomsk Standard Time": "Asia/Tomsk",
"China Standard Time": "Asia/Shanghai",
"North Asia East Standard Time": "Asia/Irkutsk",
"Singapore Standard Time": "Asia/Singapore",
"W. Australia Standard Time": "Australia/Perth",
"Taipei Standard Time": "Asia/Taipei",
"Ulaanbaatar Standard Time": "Asia/Ulaanbaatar",
"Aus Central W. Standard Time": "Australia/Eucla",
"Transbaikal Standard Time": "Asia/Chita",
"Tokyo Standard Time": "Asia/Tokyo",
"North Korea Standard Time": "Asia/Pyongyang",
"Korea Standard Time": "Asia/Seoul",
"Yakutsk Standard Time": "Asia/Yakutsk",
"Cen. Australia Standard Time": "Australia/Adelaide",
"AUS Central Standard Time": "Australia/Darwin",
"E. Australia Standard Time": "Australia/Brisbane",
"AUS Eastern Standard Time": "Australia/Sydney",
"West Pacific Standard Time": "Pacific/Port_Moresby",
"Tasmania Standard Time": "Australia/Hobart",
"Vladivostok Standard Time": "Asia/Vladivostok",
"Lord Howe Standard Time": "Australia/Lord_Howe",
"Bougainville Standard Time": "Pacific/Bougainville",
"Russia Time Zone 10": "Asia/Srednekolymsk",
"Magadan Standard Time": "Asia/Magadan",
"Norfolk Standard Time": "Pacific/Norfolk",
"Sakhalin Standard Time": "Asia/Sakhalin",
"Central Pacific Standard Time": "Pacific/Guadalcanal",
"Russia Time Zone 11": "Asia/Kamchatka",
"New Zealand Standard Time": "Pacific/Auckland",
"UTC+12": "Pacific/Wallis",
"Fiji Standard Time": "Pacific/Fiji",	
"Chatham Islands Standard Time": "Pacific/Chatham",	
"UTC+13": "Pacific/Enderbury",	
"Tonga Standard Time": "Pacific/Tongatapu",	
"Samoa Standard Time": "Pacific/Apia",	
"Line Islands Standard Time": "Pacific/Kiritimati",
"Z": "Etc/GMT"
};

// Source: src/runtime/execution.js
var scriptPrp = PropertiesService.getScriptProperties();
var userPrp = PropertiesService.getUserProperties();
var rosterUserToken = userPrp.getProperty("rosterUserToken");
var rosterUserId = userPrp.getProperty("rosterUserId");
var defaultRosterToken = "TOKEN";
var defaultRosterId = "ID";
var updateAvailable = String(
  Number(scriptPrp.getProperty("availableVersion") || version) > version
);

/** The account whose calendars and properties this execution works on. */
function currentUserEmail_() {
  return Session.getEffectiveUser().getEmail().trim().toLowerCase();
}

/** The signed-in visitor. Differs from the account above if a deployment runs as its owner. */
function visitorEmail_() {
  return Session.getActiveUser().getEmail().trim().toLowerCase();
}

function appOwnerEmail_() {
  return scriptPrp.getProperty("appOwnerEmail");
}

/**
 * Every log line of the app names the account, so the runs of several people can be
 * told apart in the shared execution log. The part before the @ is enough to recognise
 * someone without repeating whole addresses in every line.
 */
var logAccount = null;
function log_(message) {
  if (logAccount === null) {
    try {
      logAccount = currentUserEmail_().split("@")[0] || "?";
    } catch (err) {
      logAccount = "?";
    }
  }
  Logger.log("[" + logAccount + "] " + message);
}

/** One log line with the time since `started`, to see where a slow run spends it. */
function logDuration_(label, started) {
  log_(label + ": " + ((Date.now() - started) / 1000).toFixed(1) + " s");
}

function hasTrigger_(handler) {
  return ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() == handler;
  });
}

/**
 * Runs and settings changes of one account exclude each other, while other accounts
 * continue independently. Google releases the lock when the execution ends, so a
 * crashed run cannot leave it behind.
 */
function withUserLock_(busyMessage, action) {
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) throw new Error(busyMessage);
  try {
    return action();
  } finally {
    lock.releaseLock();
  }
}

/** A single property value may hold about 9 kB. */
function byteLength_(text) {
  return Utilities.newBlob(text).getBytes().length;
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text)
    .map(function (byte) {
      return ("0" + (byte & 255).toString(16)).slice(-2);
    })
    .join("");
}

/** Functions used before the preferences are loaded fall back to the default zone. */
function rosterZone_() {
  return typeof rosterTimeZone == "undefined" ? "Europe/Berlin" : rosterTimeZone;
}

var syncStats = { added: 0, updated: 0, removed: 0 };

var defaultMaxRetries = 4;

var syncStartedAt = 0;

var syncDeadline = 0;

var syncPhase = "";

var rosterFeedCache;

var rosterShiftKeys = new Map(); // shift UID -> identity (code + start), see shiftBlockToICal_
var teamRosterCache = {};

var teamCrewCache = new WeakMap();

var legendDutySpanCache = new Map();
var zoneDayOffsetCache = new Map();

var teamRosterTokenReady = false;

var lastProgressWriteAt = 0;

var lastProgressPhase = "";

function resetRosterCaches_() {
  rosterFeedCache = undefined;
  rosterShiftKeys = new Map();
  teamRosterCache = {};
  teamCrewCache = new WeakMap();
  legendDutySpanCache = new Map();
  zoneDayOffsetCache = new Map();
  teamRosterTokenReady = false;
}

function publishSyncProgress_(phase, completed, total, status) {
  if (!syncStartedAt && !status) return;
  var now = Date.now();
  if (!status && phase == lastProgressPhase && now - lastProgressWriteAt < 1000) return;
  userPrp.setProperty(
    "syncProgress",
    JSON.stringify({
      phase: phase,
      completed: completed,
      total: total,
      status: status || "running",
      updatedAt: now,
      startedAt: syncStartedAt || now
    })
  );
  lastProgressWriteAt = now;
  lastProgressPhase = phase;
}

function checkSyncBudget_() {
  if (syncDeadline && Date.now() >= syncDeadline) {
    var error = new Error("Die Synchronisierung wird beim nächsten Lauf fortgesetzt: " + syncPhase);
    error.syncPaused = true;
    throw error;
  }
}

function timedSyncStep_(label, action) {
  syncPhase = label;
  checkSyncBudget_();
  var started = Date.now();
  log_("Start: " + label);
  try {
    publishSyncProgress_(label);
    return action();
  } finally {
    log_(label + ": " + ((Date.now() - started) / 1000).toFixed(1) + " s");
  }
}

function scheduleSyncContinuation_(remaining) {
  if (!hasTrigger_("continueSync_")) {
    var continuation = ScriptApp.newTrigger("continueSync_").timeBased().after(60000).create();
    userPrp.setProperty("syncContinuationTriggerId", String(continuation.getUniqueId()));
  }
  userPrp.setProperty("syncPending", "true");
  userPrp.setProperty("nextContinuationAt", String(Date.now() + 60000));
  publishSyncProgress_(
    remaining
      ? "Alte Termine werden im nächsten Lauf weiter zugeordnet"
      : syncPhase + " – Fortsetzung geplant",
    undefined,
    undefined,
    "pending"
  );
}

function continueSync_(event) {
  return startSync(event);
}

// Source: src/runtime/failures.js
/**
 * Runs the specified function with exponential backoff and returns the result.
 * Will return null if the function did not succeed afterall.
 *
 * @param {function} func - The function that should be executed
 * @param {Number} maxRetries - How many times the function should try if it fails
 * @return {?Calendar.Event} The Calendar.Event that was added in the calendar, null if func did not complete successfully
 */
var backoffRecoverableErrors = [
  "service invoked too many times in a short time",
  "rate limit exceeded",
  "internal error"
];

function callWithBackoff_(func, maxRetries) {
  if (maxRetries === undefined) maxRetries = defaultMaxRetries;
  var tries = 0;
  var result;
  while (tries <= maxRetries) {
    tries++;
    try {
      checkSyncBudget_();
      result = func();
      return result;
    } catch (err) {
      if (err && (err.syncPaused || err.rosterAuthRejected)) throw err;
      var text = String((err && err.message) || err);
      if (
        text.includes("is not a function") ||
        !backoffRecoverableErrors.some(function (e) {
          return text.toLowerCase().includes(e);
        })
      ) {
        throw err; // The original error keeps its stack for the report.
      } else if (tries > maxRetries) {
        log_(`Error, giving up after trying ${maxRetries} times [${text}]`);
        return null;
      } else {
        log_("Error, Retrying... [" + text + "]");
        // Jitter of 0–99 ms keeps every delay below 8.1 s.
        var delay = Math.min(Math.pow(2, tries - 1) * 1000, 8000) + Math.floor(Math.random() * 100);
        if (syncDeadline && Date.now() + delay >= syncDeadline) {
          syncDeadline = Date.now();
          checkSyncBudget_();
        }
        Utilities.sleep(delay);
      }
    }
  }
  return null;
}

/** Share only program locations and known phase names, never raw error bodies. */
function buildSyncDiagnostic_(error) {
  var phases = [
    "Synchronisierung wird vorbereitet",
    "Quellen laden",
    "Kalender öffnen",
    "vorhandene Termine laden",
    "alte Termine zuordnen",
    "Termine auswerten",
    "Termine verarbeiten",
    "Termine prüfen",
    "Serientermine verarbeiten",
    "Serientermine wiederherstellen",
    "Laufstatus speichern"
  ];
  var phase =
    phases.find(function (name) {
      return syncPhase == name || syncPhase.endsWith(": " + name);
    }) || "Synchronisierung";
  var types = [
    "Error",
    "TypeError",
    "ReferenceError",
    "RangeError",
    "SyntaxError",
    "URIError",
    "EvalError"
  ];
  var type = error && types.indexOf(error.name) != -1 ? error.name : "Error";
  var lines = [
    "Dienstplan Sync – Fehlerbericht",
    "Version: " + version,
    "Zeit (UTC): " + new Date().toISOString(),
    "Schritt: " + phase,
    "Fehlertyp: " + type
  ];
  var message = String((error && error.message) || "");
  var missing =
    /Cannot read properties of (undefined|null) \(reading '(private|extendedProperties|id|data|length)'\)/.exec(
      message
    ) || /(?:Invalid argument|Ungültiges Argument|Argument too large): [A-Za-z]+/.exec(message);
  if (missing) lines.push("Fehler: " + missing[0]);
  var frames = [];
  String((error && error.stack) || "")
    .split("\n")
    .forEach(function (line) {
      var match =
        /(?:\(|\s)(Code|Helpers|Roster|WebApp|Preferences|Updates)(?:\.gs)?:(\d+)(?::(\d+))?\)?\s*$/.exec(
          line
        );
      if (match && frames.length < 12)
        frames.push(match[1] + ".gs:" + match[2] + (match[3] ? ":" + match[3] : ""));
    });
  return lines
    .concat(
      frames.length ? ["Dateipositionen:"].concat(frames) : ["Keine Dateiposition verfügbar."]
    )
    .join("\n");
}

// Source: src/runtime/alerts.js
/**
 * Email warnings. An account's own scheduled runs warn that account. Problems that
 * stop runs altogether cannot report themselves, so every account also keeps a short
 * health record in the script properties, which the owner's daily check reads.
 */
var alertDelayMs = 3600000; // A problem must last this long before the first warning,
var alertReminderMs = 3 * 86400000; // and warnings about it repeat at most this often.

/** The address for warnings: the account's own, another one, or "" when switched off. */
function alertRecipient_(setting) {
  if (setting === true) return currentUserEmail_();
  return typeof setting == "string" ? setting.trim() : "";
}

function alertTime_(value) {
  return Utilities.formatDate(new Date(value), "Europe/Berlin", "dd.MM.yyyy 'um' HH:mm 'Uhr'");
}

function appUrl_() {
  try {
    return ScriptApp.getService().getUrl() || "";
  } catch (err) {
    return "";
  }
}

/** Sends one warning; a failed delivery is logged and never hides the run's own result. */
function sendAlertMail_(recipient, subject, lines) {
  var url = appUrl_();
  try {
    MailApp.sendEmail(
      recipient,
      "Dienstplan Sync: " + subject,
      lines.concat(url ? ["", "App öffnen: " + url] : []).join("\n")
    );
    return true;
  } catch (err) {
    log_("Could not send warning email: " + ((err && err.message) || err));
    return false;
  }
}

function readJson_(text, fallback) {
  try {
    var value = JSON.parse(text);
    return value && typeof value == "object" ? value : fallback;
  } catch (err) {
    return fallback;
  }
}

/**
 * Records the outcome of a finished run and sends the warnings that are due. A problem
 * is reported once it has lasted an hour (an expired roster access at once), then at
 * most every three days, and a short note follows once a run succeeds again. Runs
 * started by hand never send mail, because the person sees the result on screen.
 *
 * @param {?Error|string} error - Why the run failed, or null for a successful run
 * @param {boolean} scheduled - Whether a trigger started the run
 */
function trackRunOutcome_(error, scheduled) {
  try {
    var now = Date.now(),
      saved = userPrp.getProperties();
    var settings = readUserPreferences_(saved);
    var recipient = alertRecipient_(settings.errorNotificationEmail);
    var alert = readJson_(saved.alertState, {});
    if (!error) {
      if (alert.sentAt && recipient)
        sendAlertMail_(recipient, "Synchronisierung läuft wieder", [
          "Hallo,",
          "",
          "die automatische Synchronisierung deines Dienstplans läuft wieder.",
          "Das Problem bestand seit " + alertTime_(alert.since) + ".",
          "",
          "Es ist nichts weiter zu tun."
        ]);
      userPrp.deleteProperty("alertState");
      alert = {};
    } else {
      var message = String((error && error.message) || error).slice(0, 300);
      if (!alert.since) alert.since = now;
      alert.problem = message;
      alert.credential = !!(error && error.rosterAuthRejected);
      var due = alert.sentAt
        ? now - alert.sentAt >= alertReminderMs
        : alert.credential || now - alert.since >= alertDelayMs;
      if (scheduled && recipient && due && sendProblemMail_(recipient, alert)) alert.sentAt = now;
      userPrp.setProperty("alertState", JSON.stringify(alert));
    }
    writeAccountHealth_({
      at: new Date(now).toISOString(),
      ok: !error,
      since: alert.since ? new Date(alert.since).toISOString() : null,
      problem: alert.problem || null,
      credential: !!alert.credential,
      enabled: hasTrigger_("startSync"),
      interval: getValidTriggerFrequency_(settings.howFrequent),
      notify: recipient
    });
  } catch (err) {
    log_("Could not record the run outcome: " + ((err && err.message) || err));
  }
}

function sendProblemMail_(recipient, alert) {
  return sendAlertMail_(
    recipient,
    alert.credential ? "Dienstplan-Zugang abgelaufen" : "Synchronisierung gestört",
    [
      "Hallo,",
      "",
      "die automatische Synchronisierung deines Dienstplans funktioniert seit " +
        alertTime_(alert.since) +
        " nicht.",
      "",
      "Meldung: " + alert.problem,
      "",
      alert.credential
        ? "Bitte öffne die App und trage über „Neuen Token eintragen“ einen neuen Dienstplan-Token ein."
        : "Deine Kalendertermine bleiben unverändert, und jeder geplante Lauf versucht es erneut. Bleibt das Problem, findest du in der App unter „Fehlerdetails“ einen Bericht zur Weitergabe.",
      "",
      "Solange das Problem besteht, kommt höchstens alle drei Tage eine Erinnerung, und eine kurze Nachricht, sobald es wieder läuft.",
      "Abschalten lassen sich diese E-Mails in der App unter Einstellungen → „E-Mail bei Fehlern“."
    ]
  );
}

/** A run that only paused or continues later still shows that the triggers work. */
function noteRunActivity_() {
  writeAccountHealth_({ at: new Date().toISOString(), enabled: true });
}

/** The shared record the owner's daily check reads. Each account writes only its own. */
function writeAccountHealth_(patch) {
  try {
    var email = currentUserEmail_();
    if (!email) return;
    var key = "health:" + email;
    var record = readJson_(scriptPrp.getProperty(key), {});
    scriptPrp.setProperty(key, JSON.stringify(Object.assign(record, patch, { email: email })));
  } catch (err) {
    log_("Could not update the account health: " + ((err && err.message) || err));
  }
}

function deleteAccountHealth_() {
  var email = currentUserEmail_();
  if (!email) return;
  scriptPrp.deleteProperty("health:" + email);
  scriptPrp.deleteProperty("stallWarned:" + email);
}

/**
 * How an account is doing, judged from its record: "paused", "stalled" (enabled but no
 * run for three intervals, at least two hours), "failing", "ok" or "unknown".
 */
function accountHealthState_(record, now) {
  if (!record || !record.enabled) return { kind: "paused" };
  var interval = Math.max(Number(record.interval) || 15, 5) * 60000;
  var lastSign = Math.max(
    Date.parse(record.at || "") || 0,
    Date.parse(record.enabledAt || "") || 0
  );
  if (lastSign && now - lastSign > Math.max(3 * interval, 7200000))
    return { kind: "stalled", since: lastSign };
  if (record.ok === false && record.since) {
    var since = Date.parse(record.since);
    if (record.credential || now - since >= alertDelayMs)
      return { kind: "failing", since: since, problem: record.problem || "" };
  }
  return { kind: lastSign ? "ok" : "unknown", at: Date.parse(record.at || "") || null };
}

/** Health of every invited account, for the owner. Records of removed accounts are dropped. */
function readAccountHealth_(now) {
  var owner = appOwnerEmail_();
  var invited = new Set(
    [owner].concat(
      invitedEmails_().map(function (value) {
        return String(value).trim().toLowerCase();
      })
    )
  );
  var all = scriptPrp.getProperties(),
    result = {};
  Object.keys(all).forEach(function (key) {
    if (key.indexOf("health:") != 0) return;
    var record = readJson_(all[key], null);
    if (!record || !invited.has(record.email)) {
      scriptPrp.deleteProperty(key);
      return;
    }
    result[record.email] = Object.assign(accountHealthState_(record, now), {
      notify: record.notify || ""
    });
  });
  return result;
}

/**
 * The owner's daily check. Accounts whose runs stopped are warned directly, because
 * they cannot notice it themselves; the owner gets one summary of all accounts with
 * problems when that list changes, and at most every three days while it stays.
 */
function reviewAccountHealth_() {
  var now = Date.now(),
    health = readAccountHealth_(now),
    problems = [];
  Object.keys(health)
    .sort()
    .forEach(function (email) {
      var state = health[email],
        warnedKey = "stallWarned:" + email;
      if (state.kind != "stalled") {
        scriptPrp.deleteProperty(warnedKey);
        if (state.kind == "failing")
          problems.push({
            email: email,
            kind: state.kind,
            since: state.since,
            text: "Fehler seit " + alertTime_(state.since) + ": " + state.problem
          });
        return;
      }
      problems.push({
        email: email,
        kind: state.kind,
        since: state.since,
        text:
          "kein automatischer Lauf seit " +
          alertTime_(state.since) +
          " – ist das Skriptprojekt mit diesem Konto als Betrachter geteilt? Ohne diese Freigabe lehnt Google automatische Läufe ab."
      });
      var warned = Date.parse(scriptPrp.getProperty(warnedKey) || "") || 0;
      if (state.notify && now - warned >= alertReminderMs) {
        var sent = sendAlertMail_(state.notify, "Automatische Synchronisierung angehalten", [
          "Hallo,",
          "",
          "für dein Konto hat seit " +
            alertTime_(state.since) +
            " kein automatischer Lauf mehr stattgefunden.",
          "Google schaltet die automatischen Läufe zum Beispiel ab, wenn eine Freigabe abgelaufen ist, und erneute Zustimmung schaltet sie nicht wieder ein.",
          "",
          "Bitte öffne die App einmal. Dabei wird die automatische Synchronisierung neu eingerichtet; falls Google nach deiner Zustimmung fragt, bestätige sie. Einstellungen und Termine bleiben erhalten.",
          "Hilft das nicht, muss der App-Inhaber das Skriptprojekt für dein Konto freigeben. Er erhält dazu ebenfalls eine Nachricht.",
          "",
          "Diese Nachricht kommt aus der täglichen Prüfung des App-Inhabers.",
          "Abschalten lassen sich diese E-Mails in der App unter Einstellungen → „E-Mail bei Fehlern“."
        ]);
        if (sent) scriptPrp.setProperty(warnedKey, new Date(now).toISOString());
      }
    });
  var signature = JSON.stringify(
    problems.map(function (problem) {
      return [problem.email, problem.kind, problem.since];
    })
  );
  var last = readJson_(scriptPrp.getProperty("healthSummary"), {});
  var owner = appOwnerEmail_();
  var due = problems.length
    ? signature != last.signature || now - (last.at || 0) >= alertReminderMs
    : !!last.signature && last.signature != "[]";
  if (!due || !owner) return;
  var sent = sendAlertMail_(
    owner,
    problems.length
      ? problems.length + (problems.length == 1 ? " Konto braucht" : " Konten brauchen") + " Hilfe"
      : "Alle Konten laufen wieder",
    problems.length
      ? ["Diese Konten synchronisieren gerade nicht:", ""]
          .concat(
            problems.map(function (problem) {
              return "• " + problem.email + " – " + problem.text;
            })
          )
          .concat([
            "",
            "Wer Warnungen nicht abgeschaltet hat, wurde selbst benachrichtigt. Den Stand aller Konten zeigt die App unter „Zugriff verwalten“."
          ])
      : ["Alle zuletzt gemeldeten Konten synchronisieren wieder."]
  );
  if (sent)
    scriptPrp.setProperty("healthSummary", JSON.stringify({ signature: signature, at: now }));
}

// Source: src/sync/schedule.js
/**
 * Takes an intended frequency in minutes and adjusts it to be the closest
 * acceptable value to use Google "everyMinutes" trigger setting (i.e. one of
 * the following values: 1, 5, 10, 15, 30).
 *
 * @param {?integer} The manually set frequency that the user intends to set.
 * @return {integer} The closest valid value to the intended frequency setting. Defaulting to 15 if no valid input is provided.
 */
function getValidTriggerFrequency_(origFrequency) {
  if (!origFrequency || origFrequency <= 0) {
    log_("No valid frequency specified. Defaulting to 15 minutes.");
    return 15;
  }

  // Limit the original frequency to 1440
  origFrequency = Math.min(origFrequency, 1440);

  var acceptableValues = [5, 10, 15, 30].concat(
    Array.from(
      {
        length: 24
      },
      (_, i) => (i + 1) * 60
    )
  ); // [5, 10, 15, 30, 60, 120, ..., 1440]

  // Find the smallest acceptable value greater than or equal to the original frequency
  var roundedUpValue = acceptableValues.find((value) => value >= origFrequency);

  // Called on every page load and run; only an actual adjustment is worth a log line.
  if (roundedUpValue != origFrequency)
    log_("Intended frequency = " + origFrequency + ", Adjusted frequency = " + roundedUpValue);
  return roundedUpValue;
}

/**
 * Removes all triggers created by this script (sync, install and update check),
 * so repeated installs can't accumulate triggers until the quota is hit.
 */
function deleteAllTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (
      ["startSync", "continueSync_", "install", "main", "checkForUpdates"].includes(
        triggers[i].getHandlerFunction()
      )
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function install() {
  requireWebUser_();
  // Triggers created without every permission would fail on each run, unnoticed.
  if (missingAuthorizationUrl_() !== null)
    throw new Error(
      "Für automatische Läufe fehlen Berechtigungen. Bitte über „Berechtigungen erteilen“ allen Berechtigungen zustimmen und danach die Automatik aktivieren."
    );
  return withUserLock_("Eine Synchronisierung läuft. Bitte danach erneut aktivieren", install_);
}

function install_() {
  loadUserPreferences_();
  if (
    userPrp.getProperty("migrationNeedsCalendarSelection") == "true" &&
    sourceCalendars.some(function (row) {
      return row.length < 4;
    })
  )
    throw new Error(
      "Bitte zuerst die Zielkalender der übernommenen Quellen auswählen und speichern"
    );
  requireSyncSources_();

  // Apps Script only returns triggers belonging to the current user.
  deleteAllTriggers_();
  userPrp.setProperty("syncPending", "false");

  // Schedule sync routine to explicitly repeat and schedule the initial sync
  var adjustedMinutes = getValidTriggerFrequency_(howFrequent);
  if (adjustedMinutes >= 60) {
    ScriptApp.newTrigger("startSync")
      .timeBased()
      .everyHours(adjustedMinutes / 60)
      .create();
  } else {
    ScriptApp.newTrigger("startSync").timeBased().everyMinutes(adjustedMinutes).create();
  }
  var initial = ScriptApp.newTrigger("startSync").timeBased().after(1000).create();
  userPrp.setProperty("initialSyncTriggerId", String(initial.getUniqueId()));
  userPrp.setProperties(
    {
      nextInitialAt: String(Date.now() + 1000),
      nextRegularAt: String(Date.now() + adjustedMinutes * 60000)
    },
    false
  );
  if (currentUserEmail_() == appOwnerEmail_())
    ScriptApp.newTrigger("checkForUpdates").timeBased().everyDays(1).create();
  writeAccountHealth_({
    enabled: true,
    enabledAt: new Date().toISOString(),
    interval: adjustedMinutes,
    notify: alertRecipient_(errorNotificationEmail)
  });
  return getAppState();
}

function uninstall() {
  return disconnectApp();
}

/** Only the owner needs a daily release check for this shared installation. */
/**
 * Automatic runs need every permission the app requests. Google lets people approve
 * only some of them; the app still opens, but every automatic run then fails before
 * any app code starts.
 *
 * @return {?string} Where to grant the missing permissions, or null when complete
 */
function missingAuthorizationUrl_() {
  try {
    if (typeof ScriptApp.getAuthorizationInfo != "function") return null;
    var info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    if (info.getAuthorizationStatus() != ScriptApp.AuthorizationStatus.REQUIRED) return null;
    return info.getAuthorizationUrl() || "";
  } catch (err) {
    log_("Could not read the authorization status: " + ((err && err.message) || err));
    return null;
  }
}

/** Opening the app asks Google again for permissions that were left out. */
function requireFullAuthorization_() {
  if (typeof ScriptApp.requireAllScopes == "function")
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
}

/**
 * Google switches off an account's time triggers when its authorization lapses, and
 * approving again does not switch them back on. The triggers stay listed, so the only
 * sign is that no run happens. Opening the app recreates them in that case.
 *
 * @return {boolean} Whether the triggers were recreated
 */
function repairStalledTriggers_() {
  try {
    if (!hasTrigger_("startSync") || missingAuthorizationUrl_() !== null) return false;
    var saved = userPrp.getProperties();
    var progress = readJson_(saved.syncProgress, null);
    if (!syncLooksStalled_(saved, true, progress, readUserPreferences_(saved).howFrequent))
      return false;
    withUserLock_("Eine Synchronisierung läuft gerade", install_);
    userPrp.setProperty("triggersRepairedAt", new Date().toISOString());
    log_("Recreated time triggers that had stopped running");
    return true;
  } catch (err) {
    // The page still opens; the notice in the overview offers the same repair by hand.
    log_("Could not recreate the time triggers: " + ((err && err.message) || err));
    return false;
  }
}

/** The owner's daily check: account warnings first, then the release version. */
function checkForUpdates() {
  requireAppOwner_();
  try {
    reviewAccountHealth_();
  } catch (err) {
    log_("Account review failed: " + ((err && err.message) || err));
  }
  var response = UrlFetchApp.fetch(
    "https://raw.githubusercontent.com/JustinBaumeyer/Google-Kalender-Sync/refs/heads/release/Code.gs",
    {
      muteHttpExceptions: true
    }
  );
  if (response.getResponseCode() != 200) throw new Error("Could not check for updates");
  var match = /^\s*var\s+version\s*=\s*(\d+)\s*;/m.exec(response.getContentText());
  if (!match) throw new Error("Update source has no version number");
  updateAvailable = String(Number(match[1]) > version);
  scriptPrp.setProperty("availableVersion", match[1]);
  scriptPrp.setProperty("updateAvailable", updateAvailable);
}

// Source: src/sync/run.js
function startSync(event) {
  requireSyncUser_(event);
  // A user lock prevents overlapping runs for this account while allowing other
  // users to sync independently. It is released automatically when
  // the execution ends, so a crashed run can never leave a stale lock behind.
  var lock = LockService.getUserLock();
  if (!lock.tryLock(0)) {
    throw new Error("Für dein Konto läuft bereits eine Synchronisierung");
  }

  // A continuation that stops at the same step without writing anything cannot finish
  // that step either; it would repeat every minute and use up Google's daily limit.
  var pausedBefore = null,
    statsAtStart = "";
  try {
    syncStartedAt = Date.now();
    syncPhase = "Synchronisierung wird vorbereitet";
    syncDeadline = syncStartedAt + 240000; // Leave room before Google's execution limit.
    lastProgressWriteAt = 0;
    lastProgressPhase = "";
    syncStats =
      userPrp.getProperty("syncPending") == "true"
        ? JSON.parse(userPrp.getProperty("syncStats") || '{"added":0,"updated":0,"removed":0}')
        : { added: 0, updated: 0, removed: 0 };
    if (userPrp.getProperty("syncPending") == "true")
      pausedBefore = userPrp.getProperty("lastSyncPhase");
    statsAtStart = JSON.stringify(syncStats);
    userPrp.deleteProperty("upcomingEvents"); // Remove the former dashboard preview cache.
    publishSyncProgress_("Synchronisierung wird vorbereitet", undefined, undefined, "running");
    if (event && event.triggerUid) {
      var initialId = userPrp.getProperty("initialSyncTriggerId");
      var continuationId = userPrp.getProperty("syncContinuationTriggerId");
      if (event.triggerUid == initialId) userPrp.deleteProperty("nextInitialAt");
      else if (event.triggerUid == continuationId) userPrp.deleteProperty("nextContinuationAt");
      else
        userPrp.setProperty(
          "nextRegularAt",
          String(Date.now() + getValidTriggerFrequency_(readUserPreferences_().howFrequent) * 60000)
        );
      ["initialSyncTriggerId", "syncContinuationTriggerId"].forEach(function (key) {
        if (userPrp.getProperty(key) == event.triggerUid) userPrp.deleteProperty(key);
      });
      ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (
          trigger.getHandlerFunction() == "continueSync_" &&
          trigger.getUniqueId() == event.triggerUid
        )
          ScriptApp.deleteTrigger(trigger);
      });
    }
    loadUserPreferences_();
    requireSyncSources_();
    var migrationRemaining = runSync_() || 0;
    timedSyncStep_("Laufstatus speichern", function () {
      userPrp.setProperties(
        {
          lastSyncAt: new Date().toISOString(),
          lastSyncError: "",
          lastSyncDurationMs: String(Date.now() - syncStartedAt),
          migrationRemaining: String(migrationRemaining),
          syncPending: "false",
          syncStats: JSON.stringify(syncStats)
        },
        false
      );
      userPrp.deleteProperty("lastSyncDiagnostic");
    });
    if (migrationRemaining) {
      scheduleSyncContinuation_(migrationRemaining);
      noteRunActivity_();
    } else {
      if (
        isRosterSourceActive_() &&
        sourceCalendars.some(function (row) {
          return row[0] == "ROSTER";
        })
      )
        markRosterConnectionVerified_();
      userPrp.setProperties(
        { lastSyncStats: JSON.stringify(syncStats), syncHistory: recordSyncRun_() },
        false
      );
      userPrp.deleteProperty("rosterAuthRejected");
      publishSyncProgress_("Synchronisierung abgeschlossen", 1, 1, "complete");
      trackRunOutcome_(null, !!(event && event.triggerUid));
    }
    if (!event || !event.triggerUid) return getAppState();
  } catch (err) {
    var noProgress =
      err.syncPaused &&
      pausedBefore == syncPhase &&
      JSON.stringify(syncStats) == statsAtStart &&
      !/alte Termine zuordnen$/.test(syncPhase);
    if (noProgress)
      err = new Error(
        "„" +
          syncPhase +
          "“ dauert länger, als ein Lauf erlaubt, auch nach einer Fortsetzung. Der nächste geplante Lauf versucht es erneut."
      );
    if (err.syncPaused) {
      log_(err.message);
      userPrp.setProperties(
        { lastSyncError: "", lastSyncPhase: syncPhase, syncStats: JSON.stringify(syncStats) },
        false
      );
      scheduleSyncContinuation_();
      noteRunActivity_();
      if (!event || !event.triggerUid) return getAppState();
      return;
    }
    var message = err.message || err;
    log_("Sync failed: " + message);
    if (err.rosterAuthRejected) {
      userPrp.setProperty("rosterAuthRejected", "true");
      userPrp.deleteProperty("connectionKey"); // The overview no longer claims a working connection.
    }
    userPrp.setProperties(
      {
        lastSyncDiagnostic: buildSyncDiagnostic_(err),
        lastSyncError: String(message),
        syncPending: "false",
        syncHistory: recordSyncRun_(message)
      },
      false
    );
    publishSyncProgress_(String(message), undefined, undefined, "failed");
    trackRunOutcome_(err, !!(event && event.triggerUid));
    throw err;
  } finally {
    syncDeadline = 0;
    lock.releaseLock();
  }
}

// Resume at the interrupted target; each target owns its loading and write phases.
function runSync_() {
  resetRosterCaches_();
  if (onlyFutureEvents) startUpdateTime = new ICAL.Time.fromJSDate(new Date());
  var calendars = condenseCalendarMap_(sourceCalendars);
  var calendarLookup = {};
  var failedCalendars = [];
  var firstFailure = null;
  var migration = {
    enabled: userPrp.getProperty("adoptLegacyEvents") == "true",
    adopted: 0,
    remaining: 0,
    deferred: false
  };
  var resumeAt = Number(userPrp.getProperty("syncResumeCalendar")) || 0;
  for (var position = 0; position < calendars.length; position++) {
    var calendarIndex = (resumeAt + position) % calendars.length;
    var calendar = calendars[calendarIndex];
    userPrp.setProperty("syncResumeCalendar", String(calendarIndex));
    try {
      if (!syncTargetCalendar_(calendar, migration, calendarLookup))
        failedCalendars.push(calendar[0]);
    } catch (err) {
      if (err && err.syncPaused) throw err;
      // Targets are independent. One broken source or calendar must not keep the
      // others from syncing; the first error is still reported for this run.
      log_("Sync of " + calendar[0] + " failed: " + ((err && err.message) || err));
      if (!firstFailure) firstFailure = { error: err, phase: syncPhase };
    }
  }
  if (firstFailure || failedCalendars.length) userPrp.deleteProperty("syncResumeCalendar");
  if (firstFailure) {
    syncPhase = firstFailure.phase; // The error report names the step that failed.
    throw firstFailure.error;
  }
  if (failedCalendars.length)
    throw new Error(
      "Nicht alle Quellen konnten geladen werden. Unverändert bleiben: " +
        failedCalendars.join(", ")
    );
  if (migration.enabled && !migration.remaining && !migration.deferred)
    userPrp.deleteProperty("adoptLegacyEvents");
  userPrp.deleteProperty("syncResumeCalendar");
  log_(
    "Sync finished: " +
      syncStats.added +
      " added, " +
      syncStats.updated +
      " updated, " +
      syncStats.removed +
      " removed"
  );
  return migration.remaining;
}

/**
 * Takes an array of ICS calendars and target Google calendars and combines them
 *
 * @param {Array.string} calendarMap - User-defined calendar map
 * @return {Array.string} Condensed calendar map
 */
function condenseCalendarMap_(calendarMap) {
  var byTarget = new Map();
  for (var mapping of calendarMap) {
    var key = mapping[3] ? "id:" + mapping[3] : "name:" + mapping[1];
    if (!byTarget.has(key))
      byTarget.set(key, mapping.length > 3 ? [mapping[1], [], mapping[3]] : [mapping[1], []]);
    byTarget.get(key)[1].push([mapping[0], mapping[2]]);
  }
  return Array.from(byTarget.values());
}

/** A short per-account run log for the overview. Repeated identical outcomes share one entry. */
var syncHistoryLimit = 12;

function readSyncHistory_(saved) {
  try {
    var rows = JSON.parse((saved && saved.syncHistory) || "[]");
    return Array.isArray(rows) ? rows : [];
  } catch (err) {
    return [];
  }
}

function syncHistoryKind_(row) {
  if (!row) return null;
  if (!row.ok) return "error:" + row.error;
  return row.added || row.updated || row.removed ? null : "quiet";
}

function nextSyncHistory_(text, entry) {
  var rows = readSyncHistory_({ syncHistory: text });
  var latest = rows[0],
    kind = syncHistoryKind_(entry);
  if (kind && kind == syncHistoryKind_(latest)) {
    latest.count = (latest.count || 1) + 1;
    latest.since = latest.since || latest.at;
    latest.at = entry.at;
    latest.ms = entry.ms;
  } else rows.unshift(entry);
  return JSON.stringify(rows.slice(0, syncHistoryLimit));
}

function recordSyncRun_(error) {
  var entry = {
    at: new Date().toISOString(),
    ms: Date.now() - syncStartedAt,
    ok: !error
  };
  if (error) entry.error = String(error).slice(0, 200);
  else Object.assign(entry, syncStats);
  return nextSyncHistory_(userPrp.getProperty("syncHistory"), entry);
}

/**
 * Google silently disables time triggers, for example after an expired authorization.
 * No run can report that, so the overview compares the last sign of life with the interval.
 */
function syncLooksStalled_(saved, enabled, progress, howFrequent) {
  if (!enabled) return false;
  var last = Math.max(
    (progress && Number(progress.updatedAt)) || 0,
    Date.parse(saved.lastSyncAt || "") || 0,
    Number(saved.nextInitialAt) || 0
  );
  if (!last) return false;
  return Date.now() - last > Math.max(3 * getValidTriggerFrequency_(howFrequent) * 60000, 7200000);
}

// Source: src/sync/target.js
var startUpdateTime;

// Per-calendar global variables (must be reset before processing each new calendar!)
var calendarEvents = [];
var cancelledCalendarEvents = [];
var calendarEventsIds = [];
var calendarEventsIndex = new Map(); // event id -> position in calendarEvents (first occurrence)
var calendarEventsMD5Set = new Set();
var icsEventsIds = [];
var ignoredEventIds = new Set();
var recurringEvents = [];
var targetCalendarId;
var targetCalendarName;
var targetCalendarTimeZone;
var syncOwnerId;

function resetTargetCalendarState_(calendar) {
  calendarEvents = [];
  cancelledCalendarEvents = [];
  calendarEventsIds = [];
  calendarEventsIndex = new Map();
  calendarEventsMD5Set = new Set();
  icsEventsIds = [];
  ignoredEventIds = new Set();
  recurringEvents = [];

  targetCalendarName = calendar[0];
}

// A failed source or incomplete Calendar listing leaves this target untouched.
function syncTargetCalendar_(calendar, migration, calendarLookup) {
  resetTargetCalendarState_(calendar);
  var sourceCalendarURLs = calendar[1];
  var vevents;

  //------------------------ Fetch URL items ------------------------

  var fetched = timedSyncStep_(targetCalendarName + ": Quellen laden", function () {
    return fetchSourceCalendars_(sourceCalendarURLs);
  });
  var responses = fetched.responses;
  if (fetched.inactiveSources) migration.deferred = true;
  if (fetched.failedSources > 0) {
    // Proceeding with missing sources would delete all their events from the
    // target calendar, so leave this calendar untouched until the next run.
    log_(
      "Skipping " +
        targetCalendarName +
        ": " +
        fetched.failedSources +
        " source(s) could not be fetched"
    );
    return false;
  }
  if (responses.length == 0) {
    log_("Skipping " + targetCalendarName + ": no sources to sync");
    return true;
  }

  //------------------------ Get target calendar information------------------------
  var targetCalendar = timedSyncStep_(targetCalendarName + ": Kalender öffnen", function () {
    return setupTargetCalendar_(targetCalendarName, calendar[2], calendarLookup);
  });
  targetCalendarId = targetCalendar.id;
  targetCalendarTimeZone = targetCalendar.timeZone;

  //------------------------ Parse existing events --------------------------
  if (addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar) {
    if (!loadTargetCalendarEvents_(migration)) {
      log_("Skipping " + targetCalendarName + ": could not list existing events");
      return false;
    }

    //------------------------ Parse ical events --------------------------

    vevents = timedSyncStep_(targetCalendarName + ": Termine auswerten", function () {
      return consolidateVevents_(parseResponses_(responses));
    });
    log_(
      targetCalendarName + ": " + vevents.length + " events from " + responses.length + " source(s)"
    );
  }

  //------------------------ Process ical events ------------------------
  if (addEventsToCalendar || modifyExistingEvents) {
    var calendarTz =
      targetCalendarTimeZone ||
      callWithBackoff_(function () {
        return Calendar.Settings.get("timezone").value;
      });

    vevents.forEach(function (e, index) {
      syncPhase = targetCalendarName + ": Termine verarbeiten";
      checkSyncBudget_();
      publishSyncProgress_(syncPhase, index, vevents.length);
      processEvent_(e, calendarTz);
    });
    publishSyncProgress_(
      targetCalendarName + ": Termine verarbeitet",
      vevents.length,
      vevents.length
    );
  }

  //------------------------ Remove old events from calendar ------------------------
  if (addEventsToCalendar && cancelledCalendarEvents.length) {
    syncPhase = targetCalendarName + ": Serientermine wiederherstellen";
    restoreDeletedOccurrences_(vevents, targetCalendarTimeZone);
  }
  // An inactive source's events cannot be compared with the partial feed.
  if (removeEventsFromCalendar && !fetched.inactiveSources) {
    syncPhase = targetCalendarName + ": Termine prüfen";
    processEventCleanup_();
  }

  //------------------------ Add Recurring Event Instances ------------------------
  syncPhase = targetCalendarName + ": Serientermine verarbeiten";
  for (var recEvent of recurringEvents) {
    checkSyncBudget_();
    processEventInstance_(recEvent);
  }
  return true;
}

// Source: src/sync/sources.js
function hasRosterCredentials_(employee, token) {
  return (
    /^[0-9]+$/.test(String(employee || "")) &&
    typeof token == "string" &&
    !!token.trim() &&
    token.trim() != defaultRosterToken
  );
}

function isRosterSourceActive_() {
  return addRosterToCal && hasRosterCredentials_(rosterUserId, rosterUserToken);
}

function hasActiveSyncSources_() {
  return sourceCalendars.some(function (row) {
    if (teamRosterEmployeeId_(row[0])) return hasRosterCredentials_(rosterUserId, rosterUserToken);
    return row[0] != "ROSTER" || isRosterSourceActive_();
  });
}

function requireSyncSources_() {
  if (!hasActiveSyncSources_())
    throw new Error("Bitte zuerst einen Dienstplan verbinden oder ein Kalender-Abo hinzufügen.");
}

/** Failed sources prevent cleanup; inactive sources preserve their existing events. */
function fetchSourceCalendars_(sourceCalendarURLs, fetchRoster) {
  var result = [];
  var failedSources = 0;
  var inactiveSources = 0;
  for (var source of sourceCalendarURLs) {
    var url = source[0].replace("webcal://", "https://");
    var colorId = source[1];

    var colleague = teamRosterEmployeeId_(source[0]);
    if (colleague) {
      if (!hasRosterCredentials_(rosterUserId, rosterUserToken)) {
        inactiveSources++;
        continue;
      }
      var teamContent = fetchRoster
        ? fetchRoster(function () {
            return getColleagueRosterICal_(colleague, true);
          })
        : getColleagueRosterICal_(colleague);
      result.push([teamContent, colorId]);
    } else if (source[0] == "ROSTER") {
      if (!isRosterSourceActive_()) {
        inactiveSources++;
        continue;
      }
      if (rosterFeedCache === undefined) rosterFeedCache = getRosterICal_();
      var rosterContent = rosterFeedCache;
      if (rosterContent != null) result.push([rosterContent, colorId]);
      else failedSources++;
    } else {
      var urlContent = null;
      try {
        urlContent = fetchSubscription_(url);
      } catch (err) {
        if (err && err.syncPaused) throw err;
        // An unreachable subscription counts as a failed source: its target stays
        // untouched, while the other calendars continue.
        log_("Could not load subscription: " + ((err && err.message) || err));
      }

      if (urlContent != null) result.push([urlContent, colorId]);
      else failedSources++;
    }
  }

  return { responses: result, failedSources: failedSources, inactiveSources: inactiveSources };
}

/** Loads one ICS subscription; repairs the truncated files some Outlook servers send. */
function fetchSubscription_(url) {
  return callWithBackoff_(function () {
    var urlResponse = UrlFetchApp.fetch(url, {
      validateHttpsCertificates: false,
      muteHttpExceptions: true
    });
    if (urlResponse.getResponseCode() != 200)
      //Throw here to make callWithBackoff_ run again
      throw (
        // Only the host: subscription links often carry a private key.
        "HTTP " +
        urlResponse.getResponseCode() +
        " from " +
        url.replace(/^([a-z]+:\/\/[^/?#]+).*$/i, "$1")
      );

    var icsContent = urlResponse.getContentText();
    const icsRegex = RegExp("(BEGIN:VCALENDAR.*?END:VCALENDAR)", "s");
    var icsMatch = icsRegex.exec(icsContent);
    if (icsMatch == null) {
      // Microsoft Outlook has a bug that sometimes results in incorrectly formatted ics files. This tries to fix that problem.
      // Add END:VEVENT for every BEGIN:VEVENT that's missing it
      const veventRegex = /BEGIN:VEVENT(?:(?!END:VEVENT).)*?(?=.BEGIN|.END:VCALENDAR|$)/gs;
      icsContent = icsContent.replace(veventRegex, (match) => match + "\nEND:VEVENT");

      // Add END:VCALENDAR if missing
      if (!icsContent.endsWith("END:VCALENDAR")) {
        icsContent += "\nEND:VCALENDAR";
      }
      icsMatch = icsRegex.exec(icsContent);
      if (icsMatch == null) {
        log_("A subscription did not return a calendar file");
        return null;
      }
    }
    return icsMatch[0];
  });
}

// Source: src/calendar/targets.js
// Calendar IDs are kept alongside legacy source rows, so renames and duplicate
// display names cannot silently redirect an existing connection.
function readWritableCalendarRows_() {
  var result = [],
    pageToken;
  do {
    checkSyncBudget_();
    var options = { showHidden: true, minAccessRole: "writer", maxResults: 250 };
    if (pageToken) options.pageToken = pageToken;
    var page = Calendar.CalendarList.list(options);
    result = result.concat(
      (page.items || []).filter(function (cal) {
        return ["owner", "writer"].indexOf(cal.accessRole) != -1;
      })
    );
    pageToken = page.nextPageToken;
  } while (pageToken);
  return result;
}

function getCalendarChoices() {
  requireWebUser_();
  return readWritableCalendarRows_()
    .map(function (cal) {
      return {
        id: cal.id,
        name: cal.summaryOverride || cal.summary,
        timeZone: cal.timeZone,
        primary: !!cal.primary
      };
    })
    .sort(function (a, b) {
      return a.name.localeCompare(b.name, "de");
    });
}

function bindCalendarId_(name, id) {
  if (!userPrp.getProperty("preferences")) return;
  var settings = readUserPreferences_(),
    changed = false;
  settings.sourceCalendars.forEach(function (row) {
    if (row[1] == name && !row[3]) {
      if (row.length < 3) row.push(null);
      row[3] = id;
      changed = true;
    }
  });
  if (changed) {
    userPrp.setProperty("preferences", validatePreferences_(settings));
    sourceCalendars = settings.sourceCalendars;
  }
}

/**
 * Gets the user's Google Calendar with the specified name.
 * A new Calendar will be created if the user does not have a Calendar with the specified name.
 *
 * @param {string} targetCalendarName - The name of the calendar to return
 * @return {Calendar} The calendar retrieved or created
 */
function setupTargetCalendar_(targetCalendarName, calendarId, lookup) {
  if (calendarId) {
    var selected = Calendar.CalendarList.get(calendarId);
    if (["owner", "writer"].indexOf(selected.accessRole) == -1)
      throw new Error("Für den ausgewählten Kalender fehlt der Schreibzugriff");
    return selected;
  }
  // Reuse a complete list within this run. UI requests and later runs stay fresh.
  var calendars = lookup
    ? lookup.rows || (lookup.rows = readWritableCalendarRows_())
    : readWritableCalendarRows_();
  var matches = calendars.filter(function (cal) {
    return (cal.summaryOverride || cal.summary) == targetCalendarName;
  });
  if (matches.length > 1)
    throw new Error(
      "Mehrere Kalender heißen „" + targetCalendarName + "“. Bitte einen Zielkalender auswählen."
    );
  var targetCalendar = matches[0];

  if (targetCalendar == null) {
    log_("Creating Calendar: " + targetCalendarName);
    targetCalendar = Calendar.newCalendar();
    targetCalendar.summary = targetCalendarName;
    targetCalendar.timeZone = Calendar.Settings.get("timezone").value;
    targetCalendar = Calendar.Calendars.insert(targetCalendar);
    if (lookup) calendars.push(targetCalendar);
  }

  bindCalendarId_(targetCalendarName, targetCalendar.id);
  return targetCalendar;
}

// Source: src/calendar/existing-events.js
// Load and index complete pages before allowing any event writes or cleanup.
function loadTargetCalendarEvents_(migration) {
  //loop until we received all events
  var pageToken;
  var listFailed = false;
  do {
    var adoptLegacy = migration.enabled;
    var listParams = {
      showDeleted: false,
      privateExtendedProperty: adoptLegacy
        ? "fromGAS=true"
        : ["fromGAS=true", "syncOwner=" + syncOwnerId],
      maxResults: 2500
    };
    if (pageToken) listParams.pageToken = pageToken;
    var eventList = timedSyncStep_(targetCalendarName + ": vorhandene Termine laden", function () {
      return callWithBackoff_(function () {
        return Calendar.Events.list(targetCalendarId, listParams);
      });
    });
    if (eventList == null) {
      listFailed = true;
      break;
    }
    cancelledCalendarEvents = cancelledCalendarEvents.concat(
      (eventList.items || []).filter(function (item) {
        return item.status == "cancelled" && item.recurringEventId && item.originalStartTime;
      })
    );
    var ownItems = (eventList.items || []).filter(function (item) {
      return !!getOwnedSyncProperties_(item, adoptLegacy);
    });
    var legacyCount = ownItems.filter(function (item) {
      return !item.extendedProperties.private.syncOwner;
    }).length;
    var calendarAdopted = 0;
    for (var item of ownItems) {
      if (!item.extendedProperties.private.syncOwner) {
        // Tag only a bounded batch. All legacy IDs still participate in
        // matching below, so untouched items are never inserted twice.
        if (migration.adopted >= 100 || (syncDeadline && syncDeadline - Date.now() < 60000)) {
          migration.remaining++;
          continue;
        }
        syncPhase = targetCalendarName + ": alte Termine zuordnen";
        item.extendedProperties.private.syncOwner = syncOwnerId;
        var claimed = callWithBackoff_(function () {
          return Calendar.Events.patch(
            { extendedProperties: item.extendedProperties },
            targetCalendarId,
            item.id
          );
        });
        if (claimed == null) {
          listFailed = true;
          break;
        }
        migration.adopted++;
        calendarAdopted++;
        publishSyncProgress_(syncPhase, calendarAdopted, legacyCount);
      }
    }
    if (listFailed) break;
    calendarEvents = calendarEvents.concat(ownItems);
    pageToken = eventList.nextPageToken;
  } while (pageToken);

  if (listFailed) return false;

  log_("Fetched " + calendarEvents.length + " existing events from " + targetCalendarName);
  if (adoptLegacy)
    log_("Migration: " + migration.adopted + " zugeordnet, " + migration.remaining + " noch offen");
  for (var i = 0; i < calendarEvents.length; i++) {
    if (calendarEvents[i].extendedProperties != null) {
      calendarEventsIds[i] =
        calendarEvents[i].extendedProperties.private["rec-id"] ||
        calendarEvents[i].extendedProperties.private["id"];
      if (!calendarEventsIndex.has(calendarEventsIds[i]))
        calendarEventsIndex.set(calendarEventsIds[i], i);
      calendarEventsMD5Set.add(calendarEvents[i].extendedProperties.private["MD5"]);
    }
  }

  return true;
}

// A missing owner alone does not identify a legacy sync event.
function getOwnedSyncProperties_(event, allowLegacy) {
  if (!event || event.status == "cancelled") return null;
  var properties = event && event.extendedProperties && event.extendedProperties.private;
  if (!properties || properties.fromGAS !== "true" || !(properties["rec-id"] || properties.id))
    return null;
  return properties.syncOwner === syncOwnerId || (allowLegacy && !properties.syncOwner)
    ? properties
    : null;
}

/**
 * Deletes all events from the target calendar that no longer exist in the source calendars.
 * Past events are preserved unless removePastEventsFromCalendar is enabled.
 * Explicit ignore rules still take precedence over retention.
 */
function processEventCleanup_() {
  var feedIds = new Set(icsEventsIds);
  // Shifts the roster still reports. An older copy with another end time is a
  // leftover of a changed shift, not history worth keeping.
  var currentShifts = new Set(
    icsEventsIds
      .map(function (id) {
        return rosterShiftKeys.get(id);
      })
      .filter(Boolean)
  );
  var compareDate = new Date();
  var today = Utilities.formatDate(
    compareDate,
    targetCalendarTimeZone || Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  for (var i = 0; i < calendarEvents.length; i++) {
    checkSyncBudget_();
    publishSyncProgress_(targetCalendarName + ": Termine prüfen", i, calendarEvents.length);
    var currentID = calendarEventsIds[i];
    // Use the end so a shift crossing midnight remains eligible for cleanup
    // until it finishes. All-day DTEND is exclusive in the target timezone.
    var eventEnd = calendarEvents[i].end || calendarEvents[i].start;
    var isPast = eventEnd.dateTime
      ? new Date(eventEnd.dateTime) <= compareDate
      : eventEnd.date <= today;

    var ignored = ignoredEventIds.has(currentID) && !feedIds.has(currentID);
    if (ignored) {
      // Explicitly ignored events are always removed, regardless of their date.
      removeManagedEvent_(calendarEvents[i].id);
    } else if (
      !feedIds.has(currentID) && // Event is no longer in source
      calendarEvents[i].recurringEventId == null && // And it's not a recurring event
      (removePastEventsFromCalendar || !isPast || currentShifts.has(storedShiftKey_(currentID)))
    ) {
      removeManagedEvent_(calendarEvents[i].id);
    }
  }
}

function removeManagedEvent_(id) {
  calendarRequest_(function () {
    Calendar.Events.remove(targetCalendarId, id);
    // A successful DELETE has no response body. Only exhausted retries yield null.
    return true;
  }, "Ein Kalendertermin konnte nicht entfernt werden. Bitte später erneut synchronisieren.");
  syncStats.removed++;
}

// Source: src/calendar/consolidation.js
/**
 * Whether event consolidation is enabled. Defaults to on when the setting is
 * missing (e.g. a settings.gs that predates the feature). consolidateShifts is
 * honored as the pre-v4 name of the setting.
 */
function consolidationEnabled_() {
  if (typeof consolidateEvents !== "undefined") return !!consolidateEvents;
  if (typeof consolidateShifts !== "undefined") return !!consolidateShifts;
  return true;
}

/**
 * Maximum gap in milliseconds between two same-titled events that still get
 * consolidated. 0 (the default) merges only touching or overlapping events.
 */
function consolidateGapMillis_() {
  return typeof consolidateMaxGapMinutes !== "undefined" && consolidateMaxGapMinutes > 0
    ? consolidateMaxGapMinutes * 60000
    : 0;
}

/**
 * Consolidates single (non-recurring) events that share the same title and
 * touch or overlap in time (optionally within consolidateMaxGapMinutes) into
 * one event spanning the whole range — e.g. two duty blocks "13:45-19:00" and
 * "18:45-00:00" become one event "13:45-00:00". Applies to every source,
 * not just the roster. Description lines of merged events are combined
 * without repeating lines both parts share. The UIDs of merged-away events
 * are removed from icsEventsIds so their previously synced counterparts get
 * cleaned up from the target calendar.
 *
 * @param {Array.ICALComponent} vevents - All parsed events of the target calendar
 * @return {Array.ICALComponent} Events with merged-away duplicates removed
 */
function consolidateVevents_(vevents) {
  if (!consolidationEnabled_()) return vevents;

  var result = [];
  var singles = [];
  for (var comp of vevents) {
    // Only plain single events are consolidated; recurring events, their
    // exceptions, explicitly separate duty cells and incomplete events pass through untouched.
    var mergeable =
      String(comp.getFirstPropertyValue("x-dpsync-keep-separate")).toLowerCase() != "true" &&
      !comp.hasProperty("rrule") &&
      !comp.hasProperty("rdate") &&
      !comp.hasProperty("exdate") &&
      !comp.hasProperty("recurrence-id") &&
      comp.hasProperty("dtstart") &&
      comp.hasProperty("summary");
    if (!mergeable) {
      result.push(comp);
      continue;
    }
    try {
      var icalEvent = new ICAL.Event(comp);
      singles.push({ comp: comp, event: icalEvent, summary: String(icalEvent.summary || "") });
    } catch (e) {
      result.push(comp);
    }
  }

  singles.sort(function (a, b) {
    return a.event.startDate.compare(b.event.startDate);
  });

  var gapSeconds = consolidateGapMillis_() / 1000;
  var mergedCount = 0;
  // Track the newest event per title so an interleaved different event
  // doesn't break a chain of same-titled ones.
  var lastByKey = {};
  singles.forEach(function (item) {
    var key = item.summary + "|" + item.event.startDate.isDate;
    var prev = item.summary ? lastByKey[key] : null;
    if (prev && item.event.startDate.toUnixTime() <= prev.event.endDate.toUnixTime() + gapSeconds) {
      if (item.event.endDate.toUnixTime() > prev.event.endDate.toUnixTime()) {
        var newEnd = item.event.endDate;
        if (!newEnd.isDate && prev.event.startDate.zone && newEnd.zone != prev.event.startDate.zone)
          newEnd = newEnd.convertToZone(prev.event.startDate.zone);
        prev.event.endDate = newEnd;
      }
      mergeEventDescriptions_(prev.comp, item.comp);
      var uid = item.comp.getFirstPropertyValue("uid").toString();
      var uidIndex = icsEventsIds.indexOf(uid);
      if (uidIndex > -1) icsEventsIds.splice(uidIndex, 1);
      mergedCount++;
    } else {
      result.push(item.comp);
      lastByKey[key] = item;
    }
  });

  if (mergedCount > 0) log_("Consolidated " + mergedCount + " overlapping event(s)");
  return result;
}

/**
 * Appends the description lines of the source event that the target event
 * doesn't already contain, so a merged event keeps both parts' details
 * without repeating shared paragraphs.
 */
function mergeEventDescriptions_(targetComp, sourceComp) {
  var sourceDesc = sourceComp.getFirstPropertyValue("description");
  if (sourceDesc == null || String(sourceDesc) == "") return;
  var targetDesc = targetComp.getFirstPropertyValue("description");
  targetDesc = targetDesc == null ? "" : String(targetDesc);
  var existingLines = targetDesc.split("\n");
  var additions = String(sourceDesc)
    .split("\n")
    .filter(function (line) {
      return line.trim() != "" && existingLines.indexOf(line) == -1;
    });
  if (!additions.length) return;
  targetComp.updatePropertyWithValue(
    "description",
    (targetDesc ? targetDesc + "\n" : "") + additions.join("\n")
  );
}

// Source: src/calendar/ignore-rules.js
/**
 * Returns the configured ignore rules, or an empty array when the setting is
 * missing (e.g. a settings.gs that predates the feature).
 */
function getIgnoreRules_() {
  return typeof ignoreEvents !== "undefined" && Array.isArray(ignoreEvents) ? ignoreEvents : [];
}

var IGNORE_WEEKDAYS = {
  so: 1,
  sun: 1,
  sunday: 1,
  sonntag: 1,
  mo: 2,
  mon: 2,
  monday: 2,
  montag: 2,
  di: 3,
  tue: 3,
  tuesday: 3,
  dienstag: 3,
  mi: 4,
  wed: 4,
  wednesday: 4,
  mittwoch: 4,
  do: 5,
  thu: 5,
  thursday: 5,
  donnerstag: 5,
  fr: 6,
  fri: 6,
  friday: 6,
  freitag: 6,
  sa: 7,
  sat: 7,
  saturday: 7,
  samstag: 7
};

/**
 * Converts an ICAL.Time into the wall-clock weekday (1 = Sunday ... 7 = Saturday)
 * and "HH:MM" time the event shows in the calendar. UTC times are converted to
 * the script timezone; times with a TZID or floating times are used as written.
 */
function getIgnoreWallClock_(time) {
  if (time.zone && time.zone.tzid == "UTC") {
    var parts = Utilities.formatDate(
      time.toJSDate(),
      targetCalendarTimeZone || rosterTimeZone || Session.getScriptTimeZone(),
      "u|HH:mm"
    ).split("|");
    return { weekday: (parseInt(parts[0], 10) % 7) + 1, time: parts[1] };
  }
  var pad = function (n) {
    return (n < 10 ? "0" : "") + n;
  };
  return { weekday: time.dayOfWeek(), time: pad(time.hour) + ":" + pad(time.minute) };
}

/**
 * Checks whether a title matches a rule's title pattern: a RegExp is tested
 * against the title, a string matches case-insensitively as a substring.
 */
function ignoreTitleMatches_(pattern, title) {
  if (pattern instanceof RegExp) {
    pattern.lastIndex = 0; // Global/sticky rules are independent for each event.
    return pattern.test(title);
  }
  return title.toLowerCase().indexOf(String(pattern).trim().toLowerCase()) > -1;
}

/**
 * Normalizes a "H:MM" / "HH:MM" string to "HH:MM" for comparison.
 */
function normalizeIgnoreTime_(value) {
  var m = /^\s*(\d{1,2}):(\d{2})\s*$/.exec(String(value));
  return m ? (m[1].length == 1 ? "0" : "") + m[1] + ":" + m[2] : String(value).trim();
}

/**
 * Checks whether the provided event matches one of the rules in ignoreEvents.
 * A rule is either a title (string/RegExp) or an object
 * {title, weekday, start, end} where every given field has to match. For a
 * recurring series the first occurrence (DTSTART) is checked, and exceptions
 * of the series are matched by their own title/time.
 *
 * @param {ICAL.Component} event - The event to check
 * @return {boolean} Whether the event should be ignored
 */
function isIgnoredEvent_(event) {
  var rules = getIgnoreRules_();
  if (rules.length == 0 || !event.hasProperty("summary")) return false;

  var title = String(event.getFirstPropertyValue("summary") || "");
  var start = null,
    end = null;
  try {
    var icalEvent = new ICAL.Event(event);
    if (icalEvent.startDate) start = getIgnoreWallClock_(icalEvent.startDate);
    if (icalEvent.endDate) end = getIgnoreWallClock_(icalEvent.endDate);
  } catch (e) {
    // Without usable times only title-only rules can match.
  }

  return rules.some(function (rule) {
    if (rule == null) return false;
    if (typeof rule === "string" || rule instanceof RegExp) return ignoreTitleMatches_(rule, title);

    if (rule.title != null && !ignoreTitleMatches_(rule.title, title)) return false;
    if (rule.weekday != null) {
      var wd =
        typeof rule.weekday === "number"
          ? rule.weekday
          : IGNORE_WEEKDAYS[String(rule.weekday).trim().toLowerCase()];
      if (start == null || wd == null || start.weekday != wd) return false;
    }
    if (rule.start != null && (start == null || start.time != normalizeIgnoreTime_(rule.start)))
      return false;
    if (rule.end != null && (end == null || end.time != normalizeIgnoreTime_(rule.end)))
      return false;
    // An empty object would ignore everything, so require at least one criterion.
    return rule.title != null || rule.weekday != null || rule.start != null || rule.end != null;
  });
}

// Source: src/calendar/parse.js
/**
 * Parses all sources using ical.js.
 * Registers all found timezones with TimezoneService.
 * Creates an Array with all events and adds the event-ids to the provided Array.
 *
 * @param {Array.string} responses - Array with all ical sources
 * @return {Array.ICALComponent} Array with all events found
 */
function parseResponses_(responses) {
  var result = [];
  for (var itm of responses) {
    var resp = itm[0];
    var colorId = itm[1];
    var jcalData = ICAL.parse(resp);
    var component = new ICAL.Component(jcalData);

    ICAL.helpers.updateTimezones(component);
    var vtimezones = component.getAllSubcomponents("vtimezone");
    for (var tz of vtimezones) {
      ICAL.TimezoneService.register(tz);
    }

    var allEvents = component.getAllSubcomponents("vevent");
    if (colorId != undefined)
      allEvents.forEach(function (event) {
        event.addPropertyWithValue("color", colorId);
      });

    var calName =
      component.getFirstPropertyValue("x-wr-calname") || component.getFirstPropertyValue("name");
    if (calName != null)
      allEvents.forEach(function (event) {
        event.addPropertyWithValue("parentCal", calName);
      });

    result = [].concat(allEvents, result);
  }

  if (onlyFutureEvents) {
    result = result.filter(function (event) {
      try {
        if (
          event.hasProperty("recurrence-id") ||
          event.hasProperty("rrule") ||
          event.hasProperty("rdate") ||
          event.hasProperty("exdate")
        ) {
          //Keep recurrences to properly filter them later on
          return true;
        }
        var eventEnde;
        eventEnde = new ICAL.Time.fromString(
          event.getFirstPropertyValue("dtend").toString(),
          event.getFirstProperty("dtend")
        );
        return eventEnde.compare(startUpdateTime) >= 0;
      } catch (e) {
        return true;
      }
    });
  }

  if (getIgnoreRules_().length > 0) {
    result = result.filter(function (event) {
      if (!isIgnoredEvent_(event)) return true;
      // Remember the id so an already synced copy is removed during cleanup,
      // even if it lies in the past or is a recurring series.
      var uid = event.hasProperty("uid")
        ? event.getFirstPropertyValue("uid").toString()
        : // Same fallback id as below, so events without UID are matched too.
          Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, event.toString()).toString();
      ignoredEventIds.add(uid);
      return false;
    });
  }

  result.forEach(function (event) {
    if (!event.hasProperty("uid")) {
      event.updatePropertyWithValue(
        "uid",
        Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, event.toString()).toString()
      );
    }
    if (event.hasProperty("recurrence-id")) {
      let recID = new ICAL.Time.fromString(
        event.getFirstPropertyValue("recurrence-id").toString(),
        event.getFirstProperty("recurrence-id")
      );
      if (event.getFirstProperty("recurrence-id").getParameter("tzid")) {
        let tz = event.getFirstProperty("recurrence-id").getParameter("tzid").toString();
        if (tz in tzidreplace) {
          tz = tzidreplace[tz];
        }
        let recUTCOffset = getUTCOffsetMillis_(tz) / -1000;
        recID = recID.adjust(0, 0, 0, recUTCOffset).toString() + "Z";
        event.updatePropertyWithValue("recurrence-id", recID);
      }
      icsEventsIds.push(event.getFirstPropertyValue("uid").toString() + "_" + recID);
    } else {
      icsEventsIds.push(event.getFirstPropertyValue("uid").toString());
    }
  });

  return result;
}

// Source: src/calendar/events.js
/**
 * Creates a Google Calendar event and inserts it to the target calendar.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {string} calendarTz - The timezone of the target calendar
 */
function processEvent_(event, calendarTz) {
  //------------------------ Create the event object ------------------------
  var newEvent = createEvent_(event, calendarTz);
  if (newEvent == null) return;

  var eventId = newEvent.extendedProperties.private["id"];
  var index = calendarEventsIndex.has(eventId) ? calendarEventsIndex.get(eventId) : -1;

  //------------------------ Save instance overrides ------------------------
  //----------- To make sure the parent event is actually created -----------
  if (event.hasProperty("recurrence-id")) {
    recurringEvents.push(newEvent);
    return;
  }

  if (index == -1 && modifyExistingEvents) {
    index = findChangedShift_(eventId);
    if (index > -1) {
      // From here on the existing event belongs to the new UID, also for the cleanup.
      calendarEventsIds[index] = eventId;
      calendarEventsIndex.set(eventId, index);
    }
  }

  if (index > -1 && modifyExistingEvents) {
    calendarRequest_(function () {
      return Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
    }, "Ein geänderter Termin konnte nicht gespeichert werden. Bestehende Termine bleiben erhalten.");
    syncStats.updated++;
  } else if (index == -1 && addEventsToCalendar) {
    calendarRequest_(function () {
      return Calendar.Events.insert(newEvent, targetCalendarId);
    }, "Ein neuer Termin konnte nicht gespeichert werden. Bestehende Termine bleiben erhalten.");
    syncStats.added++;
  }
}

/**
 * Creates a Google Calendar Event based on the specified ICALEvent.
 * Will return null if the event has not changed since the last sync.
 * If onlyFutureEvents is set to true:
 * -It will return null if the event has already taken place.
 * -Past instances of recurring events will be removed
 *
 * @param {ICAL.Component} event - The event to process
 * @param {string} calendarTz - The timezone of the target calendar
 * @return {?Calendar.Event} The Calendar.Event that will be added to the target calendar
 */
function createEvent_(event, calendarTz) {
  event.removeProperty("dtstamp");
  var icalEvent = new ICAL.Event(event, {
    strictExceptions: true
  });
  if (onlyFutureEvents && checkSkipEvent_(event, icalEvent)) {
    return;
  }

  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    icalEvent.toString() + eventRenderSignature_(icalEvent.startDate.isDate)
  ).toString();
  if (calendarEventsMD5Set.has(digest)) {
    showSkipMessage &&
      log_("Skipping unchanged Event " + event.getFirstPropertyValue("uid").toString());
    return;
  }

  var newEvent;
  if (icalEvent.startDate.isDate) {
    //All-day event
    if (icalEvent.startDate.compare(icalEvent.endDate) == 0) {
      //Adjust dtend in case dtstart equals dtend as this is not valid for allday events
      icalEvent.endDate = icalEvent.endDate.adjust(1, 0, 0, 0);
    }

    newEvent = {
      start: {
        date: icalEvent.startDate.toString()
      },
      end: {
        date: icalEvent.endDate.toString()
      }
    };
  } else {
    //Normal (not all-day) event
    var tzid = icalEvent.startDate.timezone;
    if (tzids.indexOf(tzid) == -1) {
      // Windows names map to IANA ones; anything else is floating time.
      tzid = tzid in tzidreplace ? tzidreplace[tzid] : calendarTz;
    }

    newEvent = {
      start: {
        dateTime: icalEvent.startDate.toString(),
        timeZone: tzid
      },
      end: {
        dateTime: icalEvent.endDate.toString(),
        timeZone: tzid
      }
    };
  }

  if (
    event.hasProperty("url") &&
    event.getFirstPropertyValue("url").toString().substring(0, 4) == "http"
  ) {
    newEvent.source = {
      url: event.getFirstPropertyValue("url").toString(),
      title: "link"
    };
  }

  if (event.hasProperty("summary")) newEvent.summary = icalEvent.summary;

  if (addCalToTitle && event.hasProperty("parentCal")) {
    var calName = event.getFirstPropertyValue("parentCal");
    newEvent.summary = "(" + calName + ") " + newEvent.summary;
  }

  if (event.hasProperty("description")) newEvent.description = icalEvent.description;

  if (event.hasProperty("location")) newEvent.location = icalEvent.location;

  var visibility =
    eventVisibility_(overrideVisibility) ||
    (event.hasProperty("class") && eventVisibility_(event.getFirstPropertyValue("class")));
  if (visibility) newEvent.visibility = visibility;

  if (event.hasProperty("transp")) {
    var transpValue = event.getFirstPropertyValue("transp").toString().toLowerCase();
    if (transpValue == "transparent" || transpValue == "opaque")
      newEvent.transparency = transpValue;
  }

  if (icalEvent.startDate.isDate) {
    if (0 <= defaultAllDayReminder && defaultAllDayReminder <= 40320) {
      newEvent.reminders = {
        useDefault: false,
        overrides: [
          {
            method: "popup",
            minutes: defaultAllDayReminder
          }
        ]
      }; //reminder as defined by the user
    } else {
      newEvent.reminders = {
        useDefault: false,
        overrides: []
      }; //no reminder
    }
  } else {
    newEvent.reminders = {
      useDefault: true,
      overrides: []
    }; //will set the default reminders as set at calendar.google.com
  }

  if (icalEvent.isRecurring()) {
    newEvent.recurrence = parseRecurrenceRule_(event, getUTCOffsetMillis_(calendarTz));
  }

  newEvent.extendedProperties = {
    private: {
      MD5: digest,
      fromGAS: "true",
      syncOwner: syncOwnerId,
      id: icalEvent.uid
    }
  };

  if (event.hasProperty("recurrence-id")) {
    newEvent.recurringEventId = event.getFirstPropertyValue("recurrence-id").toString();
    newEvent.extendedProperties.private["rec-id"] =
      newEvent.extendedProperties.private["id"] + "_" + newEvent.recurringEventId;
  }

  if (event.hasProperty("color")) {
    let colorID = event.getFirstPropertyValue("color").toString();
    if (Object.keys(CalendarApp.EventColor).includes(colorID)) {
      newEvent.colorId = CalendarApp.EventColor[colorID];
    } else if (Object.values(CalendarApp.EventColor).includes(colorID)) {
      newEvent.colorId = colorID;
    } //else unsupported value
  }

  return newEvent;
}

/**
 * The identity (code + start) inside the UID of a stored roster shift, or null for
 * other events. Calendars hold two formats: "<code><start><end>", and
 * "roster-shift-<encoded code>-<start>" written by versions 20 to 22. Both mean the
 * same shift, so events of either format are continued instead of added again.
 */
function storedShiftKey_(id) {
  id = String(id || "");
  var match = /^(.+\d{8}T\d{6}Z)\d{8}T\d{6}Z$/.exec(id);
  if (match) return match[1];
  match = /^roster-shift-(.*)-(\d{8}T\d{6}Z)$/.exec(id);
  if (!match) return null;
  try {
    return decodeURIComponent(match[1]) + match[2];
  } catch (err) {
    return null;
  }
}

/**
 * The existing event of a roster shift whose end time changed, for example through
 * overtime. Its UID differs, but it is the same shift and keeps its calendar event.
 *
 * @param {string} uid - UID of the shift as the roster reports it now
 * @return {number} Position in calendarEvents, or -1
 */
function findChangedShift_(uid) {
  var key = rosterShiftKeys.get(uid);
  if (!key) return -1;
  var current = new Set(icsEventsIds);
  return calendarEvents.findIndex(function (event, i) {
    return (
      !current.has(calendarEventsIds[i]) &&
      event.recurringEventId == null &&
      storedShiftKey_(calendarEventsIds[i]) == key
    );
  });
}

/** Keeps only the matching values of a multi-value property and drops it once empty. */
function keepPropertyValues_(event, name, keep) {
  event.getAllProperties(name).forEach(function (property) {
    var values = property.getValues().filter(function (value) {
      return keep(value, property);
    });
    if (values.length == 0) event.removeProperty(property);
    else if (values.length == 1) property.setValue(values[0]);
    else property.setValues(values);
  });
}

/** A visibility Google Calendar accepts, or "" for anything else. */
function eventVisibility_(value) {
  var visibility = String(value == null ? "" : value).toLowerCase();
  return ["default", "public", "private", "confidential"].indexOf(visibility) != -1
    ? visibility
    : "";
}

/** A Calendar request that still fails after all retries aborts the target before any cleanup. */
function calendarRequest_(request, failureMessage) {
  var result = callWithBackoff_(request);
  if (result == null) throw new Error(failureMessage);
  return result;
}

/**
 * Settings that change how an unchanged source event is written. Mixing them into
 * the digest lets a changed setting reach events that were imported earlier.
 * Program defaults add nothing, so digests stored by earlier versions stay valid.
 *
 * @param {boolean} allDay - Whether the event is an all-day event
 * @return {string} Suffix for the change digest
 */
function eventRenderSignature_(allDay) {
  var parts = [];
  if (eventVisibility_(overrideVisibility))
    parts.push("visibility=" + eventVisibility_(overrideVisibility));
  if (addCalToTitle) parts.push("calendarInTitle");
  if (allDay && defaultAllDayReminder >= 0 && defaultAllDayReminder <= 40320)
    parts.push("allDayReminder=" + defaultAllDayReminder);
  return parts.length ? "\n#" + parts.join(";") : "";
}

/**
 * Checks if the provided event has taken place in the past.
 * Removes all past instances of the provided icalEvent object.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {ICAL.Event} icalEvent - The event to process as ICAL.Event object
 * @return {boolean} Wether it's a past event or not
 */
function checkSkipEvent_(event, icalEvent) {
  if (icalEvent.isRecurrenceException()) {
    if (
      icalEvent.startDate.compare(startUpdateTime) < 0 &&
      icalEvent.recurrenceId.compare(startUpdateTime) < 0
    ) {
      return true;
    }
  } else if (icalEvent.isRecurring()) {
    var skip = false; //Indicates if the recurring event and all its instances are in the past
    if (icalEvent.endDate.compare(startUpdateTime) < 0) {
      //Parenting recurring event is in the past
      var dtstart = event.getFirstPropertyValue("dtstart");
      var expand = new ICAL.RecurExpansion({
        component: event,
        dtstart: dtstart
      });
      var next;
      var newStartDate;
      var countskipped = 0;
      while ((next = expand.next())) {
        var diff = next.subtractDate(icalEvent.startDate);
        var tempEnd = icalEvent.endDate.clone();
        tempEnd.addDuration(diff);
        if (tempEnd.compare(startUpdateTime) < 0) {
          countskipped++;
          continue;
        }

        newStartDate = next;
        break;
      }

      if (newStartDate != null) {
        //At least one instance is in the future
        newStartDate.timezone = icalEvent.startDate.timezone;
        var diff = newStartDate.subtractDate(icalEvent.startDate);
        icalEvent.endDate.addDuration(diff);
        var newEndDate = icalEvent.endDate;
        icalEvent.endDate = newEndDate;
        icalEvent.startDate = newStartDate;

        var rrule = event.getFirstProperty("rrule");
        var recur = rrule.getFirstValue();
        if (recur.isByCount()) {
          recur.count -= countskipped;
          rrule.setValue(recur);
        }

        keepPropertyValues_(event, "exdate", function (value) {
          return new ICAL.Time.fromString(value.toString()) > newStartDate;
        });
        keepPropertyValues_(event, "rdate", function (value, property) {
          var time = new ICAL.Time.fromString(value.toString(), property);
          return time.compare(startUpdateTime) >= 0 && time.compare(icalEvent.startDate) > 0;
        });
      } else {
        //All instances are in the past
        skip = true;
      }
    }

    //Check and filter recurrence-exceptions
    for (var i = 0; i < icalEvent.except.length; i++) {
      //Exclude the instance if it was moved from future to past
      if (
        icalEvent.except[i].startDate.compare(startUpdateTime) < 0 &&
        icalEvent.except[i].recurrenceId.compare(startUpdateTime) >= 0
      ) {
        icalEvent.component.addPropertyWithValue(
          "exdate",
          icalEvent.except[i].recurrenceId.toString()
        );
      } //Re-add the instance if it is moved from past to future
      else if (
        icalEvent.except[i].startDate.compare(startUpdateTime) >= 0 &&
        icalEvent.except[i].recurrenceId.compare(startUpdateTime) < 0
      ) {
        icalEvent.component.addPropertyWithValue(
          "rdate",
          icalEvent.except[i].recurrenceId.toString()
        );
        skip = false;
      }
    }

    if (skip) {
      //Completely remove the event as all instances of it are in the past
      icsEventsIds.splice(icsEventsIds.indexOf(event.getFirstPropertyValue("uid").toString()), 1);
      return true;
    }
  } else {
    //normal events
    if (icalEvent.endDate.compare(startUpdateTime) < 0) {
      icsEventsIds.splice(icsEventsIds.indexOf(event.getFirstPropertyValue("uid").toString()), 1);
      return true;
    }
  }
  return false;
}

// Source: src/calendar/recurrence.js
/**
 * Parses the provided ICAL.Component to find all recurrence rules.
 *
 * @param {ICAL.Component} vevent - The event to parse
 * @param {number} utcOffset - utc offset of the target calendar
 * @return {Array.String} Array with all recurrence components found in the provided event
 */
function parseRecurrenceRule_(vevent, utcOffset) {
  var recurrenceRules = vevent.getAllProperties("rrule");
  var exRules = vevent.getAllProperties("exrule"); //deprecated, for compatibility only
  var exDates = vevent.getAllProperties("exdate");
  var rDates = vevent.getAllProperties("rdate");

  var recurrence = [];
  for (var recRule of recurrenceRules) {
    var recIcal = recRule.toICALString();
    var adjustedTime;

    var untilMatch = RegExp(
      "(.*)(UNTIL=)(\\d\\d\\d\\d)(\\d\\d)(\\d\\d)T(\\d\\d)(\\d\\d)(\\d\\d)(;.*|\\b)",
      "g"
    ).exec(recIcal);
    if (untilMatch != null) {
      adjustedTime = new Date(
        Date.UTC(
          parseInt(untilMatch[3], 10),
          parseInt(untilMatch[4], 10) - 1,
          parseInt(untilMatch[5], 10),
          parseInt(untilMatch[6], 10),
          parseInt(untilMatch[7], 10),
          parseInt(untilMatch[8], 10)
        )
      );
      adjustedTime = Utilities.formatDate(
        new Date(adjustedTime - utcOffset),
        "etc/GMT",
        "YYYYMMdd'T'HHmmss'Z'"
      );
      recIcal = untilMatch[1] + untilMatch[2] + adjustedTime + untilMatch[9];
    }

    recurrence.push(recIcal);
  }

  for (var exRule of exRules) {
    recurrence.push(exRule.toICALString());
  }

  for (var exDate of exDates) {
    recurrence.push(exDate.toICALString());
  }

  for (var rDate of rDates) {
    recurrence.push(rDate.toICALString());
  }

  return recurrence;
}

/**
 * Returns the current UTC offset of the given timezone in milliseconds.
 * Cached per timezone for the duration of the execution, as the lookup via
 * Utilities.formatDate is expensive when done once per event.
 *
 * @param {string} tz - IANA timezone name
 * @return {number} UTC offset in milliseconds
 */
var tzOffsetCache = {};
function getUTCOffsetMillis_(tz) {
  if (tz in tzOffsetCache) return tzOffsetCache[tz];
  var jsTime = new Date();
  var utcTime = new Date(Utilities.formatDate(jsTime, "Etc/GMT", "HH:mm:ss MM/dd/yyyy"));
  var tgtTime = new Date(Utilities.formatDate(jsTime, tz, "HH:mm:ss MM/dd/yyyy"));
  tzOffsetCache[tz] = tgtTime - utcTime;
  return tzOffsetCache[tz];
}

/**
 * Patches an existing event instance with the provided Calendar.Event.
 * The instance that needs to be updated is identified by the recurrence-id of the provided event.
 *
 * @param {Calendar.Event} recEvent - The event instance to process
 */
function processEventInstance_(recEvent) {
  var eventInstanceToPatch = findOwnedEventInstances_({
    singleEvents: true,
    privateExtendedProperty: [
      "fromGAS=true",
      "rec-id=" + recEvent.extendedProperties.private["id"] + "_" + recEvent.recurringEventId
    ]
  });

  if (eventInstanceToPatch.length == 0) {
    if (recEvent.recurringEventId.length == 10) {
      recEvent.recurringEventId += "T00:00:00Z";
    } else if (recEvent.recurringEventId.substr(-1) !== "Z") {
      recEvent.recurringEventId += "Z";
    }
    eventInstanceToPatch = findOwnedEventInstances_({
      singleEvents: true,
      orderBy: "startTime",
      timeMin: recEvent.recurringEventId,
      privateExtendedProperty: ["fromGAS=true", "id=" + recEvent.extendedProperties.private["id"]]
    });
  }

  if (eventInstanceToPatch.length == 1) {
    if (modifyExistingEvents) {
      calendarRequest_(function () {
        return Calendar.Events.update(recEvent, targetCalendarId, eventInstanceToPatch[0].id);
      }, "Ein Serientermin konnte nicht aktualisiert werden.");
      syncStats.updated++;
    }
  } else {
    if (addEventsToCalendar) {
      calendarRequest_(function () {
        return Calendar.Events.insert(recEvent, targetCalendarId);
      }, "Ein Serientermin konnte nicht gespeichert werden.");
      syncStats.added++;
    }
  }
}

/** Restore destination-only cancellations whose original date still belongs to the source series. */
function restoreDeletedOccurrences_(components, calendarTz) {
  var parents = new Map();
  calendarEvents.forEach(function (event) {
    if (event.recurrence && !event.recurringEventId) parents.set(event.id, event);
  });
  cancelledCalendarEvents.forEach(function (cancelled) {
    checkSyncBudget_();
    var parent = parents.get(cancelled.recurringEventId);
    if (!parent) return;
    var uid = parent.extendedProperties.private.id;
    var component = components.find(function (item) {
      return !item.hasProperty("recurrence-id") && String(item.getFirstPropertyValue("uid")) == uid;
    });
    if (
      !component ||
      String(component.getFirstPropertyValue("status")).toUpperCase() == "CANCELLED"
    )
      return;
    var exceptions = components.filter(function (item) {
      return item.hasProperty("recurrence-id") && String(item.getFirstPropertyValue("uid")) == uid;
    });
    var source = new ICAL.Event(component, { strictExceptions: true, exceptions: exceptions });
    if (!source.isRecurring()) return;
    var original = cancelled.originalStartTime;
    function instant(time) {
      if (time.isDate) return time.toString();
      if (time.zone && time.zone.tzid == "UTC") return time.toJSDate().getTime();
      var zone = component.getFirstProperty("dtstart").getParameter("tzid") || calendarTz;
      zone = tzidreplace[zone] || zone;
      if (tzids.indexOf(zone) == -1) zone = calendarTz;
      return Utilities.parseDate(time.toString(), zone, "yyyy-MM-dd'T'HH:mm:ss").getTime();
    }
    var target = original.date || new Date(original.dateTime).getTime();
    var iterator = source.iterator(),
      occurrence;
    for (var n = 0; n < 10000 && (occurrence = iterator.next()); n++) {
      if (n % 100 == 0) checkSyncBudget_();
      if (!!original.date != occurrence.isDate) return;
      var value = instant(occurrence);
      if (value > target) return;
      if (value != target) continue;
      var exception = source.exceptions[occurrence.toString()];
      if (
        exception &&
        String(exception.component.getFirstPropertyValue("status")).toUpperCase() == "CANCELLED"
      )
        return;
      var details = source.getOccurrenceDetails(occurrence);
      var copy = new ICAL.Component(JSON.parse(JSON.stringify(details.item.component.toJSON())));
      ["rrule", "rdate", "exdate", "recurrence-id", "duration"].forEach(function (name) {
        copy.removeAllProperties(name);
      });
      copy.updatePropertyWithValue("dtstart", details.startDate.clone());
      copy.updatePropertyWithValue("dtend", details.endDate.clone());
      copy.addPropertyWithValue("recurrence-id", occurrence.clone());
      var restored = createEvent_(copy, calendarTz);
      if (!restored) return;
      delete restored.recurringEventId;
      restored.status = "confirmed";
      calendarRequest_(function () {
        return Calendar.Events.patch(restored, targetCalendarId, cancelled.id);
      }, "Ein gelöschter Serientermin konnte nicht wiederhergestellt werden.");
      syncStats.added++;
      return;
    }
  });
}

// Pagination plus an ownership check also protects migrated recurrence overrides
// when different accounts share a destination calendar and source UID.
function findOwnedEventInstances_(options) {
  var adoptLegacy = userPrp.getProperty("adoptLegacyEvents") == "true";
  var params = Object.assign({}, options, { maxResults: 25 });
  params.privateExtendedProperty = options.privateExtendedProperty.slice();
  if (!adoptLegacy) params.privateExtendedProperty.push("syncOwner=" + syncOwnerId);
  do {
    var page = calendarRequest_(function () {
      return Calendar.Events.list(targetCalendarId, params);
    }, "Serientermine konnten nicht vollständig geladen werden.");
    var own = (page.items || []).filter(function (item) {
      return !!getOwnedSyncProperties_(item, adoptLegacy);
    });
    if (own.length) return own.slice(0, 1);
    params.pageToken = page.nextPageToken;
  } while (params.pageToken);
  return [];
}

// Source: src/roster/client.js
/** Options shared by single and parallel roster requests. */
function rosterRequestOptions_(method, payload) {
  var options = {
    validateHttpsCertificates: false,
    muteHttpExceptions: true,
    method: method,
    headers: {
      authorization: "Bearer " + rosterUserToken,
      "content-type": "application/json"
    }
  };
  if (payload != null) options.payload = payload;
  return options;
}

/**
 * Performs an authenticated request against the roster API and returns the response body.
 * Throws on any non-200 status so callWithBackoff_ can retry.
 *
 * @param {string} endpoint - Path appended to rosterUrl (e.g. "auth/refreshToken").
 * @param {string} method - HTTP method ("GET" or "POST").
 * @param {string} [payload] - Optional request body.
 * @return {string} The response body text.
 */
function rosterFetch_(endpoint, method, payload) {
  checkSyncBudget_();
  var response = UrlFetchApp.fetch(rosterUrl + endpoint, rosterRequestOptions_(method, payload));
  var status = response.getResponseCode();
  if (status == 401 || status == 403) {
    var error = new Error(
      "Dienstplan-Zugriff abgelehnt (HTTP " +
        status +
        "). Bitte Mitarbeiter-ID, Dienstplan-Token und Dienstplan-Adresse prüfen."
    );
    error.rosterAuthRejected = true;
    throw error;
  }
  if (response.getResponseCode() != 200)
    //Throw here to make callWithBackoff_ run again
    throw (
      "Error: Encountered HTTP error " + response.getContentText() + " when accessing " + endpoint
    );
  return response.getContentText();
}

/**
 * Performs several authenticated roster API requests in parallel via
 * UrlFetchApp.fetchAll (one round-trip instead of one per request).
 * By default failed requests are logged and returned as null. Strict calendar
 * sources throw instead, so a partial response cannot trigger event cleanup.
 *
 * @param {Array.<{endpoint: string, method: string, payload: ?string}>} calls - Requests to perform.
 * @param {boolean} [strict] - Reject incomplete results and mark authorization failures.
 * @return {Array.<?string>} One response body per request, null where the request failed.
 */
function rosterFetchAll_(calls, strict) {
  checkSyncBudget_();
  if (!calls.length) return [];
  var requests = calls.map(function (call) {
    return Object.assign(
      { url: rosterUrl + call.endpoint },
      rosterRequestOptions_(call.method, call.payload)
    );
  });
  return UrlFetchApp.fetchAll(requests).map(function (response, i) {
    if (response.getResponseCode() != 200) {
      if (strict) {
        var error = new Error(
          "Die Teamübersicht konnte nicht vollständig geladen werden (HTTP " +
            response.getResponseCode() +
            "). Bitte deinen Dienstplan-Zugang prüfen."
        );
        error.rosterAuthRejected = [401, 403].indexOf(response.getResponseCode()) != -1;
        throw error;
      }
      log_(
        "Roster API call to " +
          calls[i].endpoint +
          " failed with HTTP " +
          response.getResponseCode()
      );
      return null;
    }
    return response.getContentText();
  });
}

/**
 * Exchanges the saved roster token for a renewed one. The roster can accept the
 * request and still send no token, for example once the session has expired. Such
 * an answer must not replace the saved token; the person has to enter a new one.
 */
function refreshRosterToken_() {
  var body = rosterFetch_("auth/refreshToken", "GET"),
    renewed = null;
  try {
    renewed = JSON.parse(body).data;
  } catch (err) {
    // An unreadable answer is treated like a missing token below.
  }
  if (typeof renewed != "string" || !renewed.trim() || byteLength_(renewed.trim()) > 9000) {
    var error = new Error(
      "Der Dienstplan hat den Zugang nicht erneuert. Bitte in den Einstellungen unter „Zugangsdaten ändern“ einen neuen Dienstplan-Token eintragen."
    );
    error.rosterAuthRejected = true;
    throw error;
  }
  rosterUserToken = renewed.trim();
  userPrp.setProperty("rosterUserToken", rosterUserToken);
}

// Source: src/roster/ical.js
/**
 * Escapes a value for use in an iCal TEXT field per RFC 5545
 * (backslash, semicolon, comma and newlines must be escaped).
 *
 * @param {*} text - The value to escape (null/undefined become "").
 * @return {string} The escaped text.
 */
function icalEscape_(text) {
  if (text == null) return "";
  return String(text)
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\r?\n/g, "\\n");
}

/** Shared VEVENT envelope for timed and date-only roster entries. */
function generateICalEvent_(uid, startLine, endLine, summary, description, options) {
  options = options || {};
  return (
    "BEGIN:VEVENT\nUID:" +
    uid +
    "\nSEQUENCE:0\nDTSTAMP:" +
    toICalDate_(new Date()) +
    "\n" +
    startLine +
    "\n" +
    endLine +
    (options.transparent ? "\nTRANSP:TRANSPARENT" : "") +
    (options.keepSeparate ? "\nX-DPSYNC-KEEP-SEPARATE:TRUE" : "") +
    "\nSUMMARY:" +
    icalEscape_(summary) +
    "\nDESCRIPTION:" +
    icalEscape_(description) +
    "\nEND:VEVENT\n"
  );
}

/** start/end are pre-formatted iCal UTC timestamps. */
function generateICalEntry_(uid, start, end, summary, description, options) {
  return generateICalEvent_(uid, "DTSTART:" + start, "DTEND:" + end, summary, description, options);
}

/**
 * Formats the span between two dates as a roster-style duration (e.g. "08h30", "24h00").
 */
function formatDuration_(startDate, endDate) {
  var totalMinutes = Math.round((endDate - startDate) / 60000);
  var hours = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  return String(hours).padStart(2, "0") + "h" + String(minutes).padStart(2, "0");
}

/**
 * Builds an all-day VEVENT. endDate is exclusive per RFC 5545, so pass the
 * day after the last day that should be covered.
 *
 * @param {string} uid - Unique identifier for the event.
 * @param {Date} startDate - First day of the event.
 * @param {Date} endDate - Day after the last day of the event (exclusive).
 * @param {Object} [options] - transparent marks free time; keepSeparate prevents consolidation.
 */
function generateAllDayICalEntry_(uid, startDate, endDate, summary, description, options) {
  return generateICalEvent_(
    uid,
    "DTSTART;VALUE=DATE:" + Utilities.formatDate(startDate, "GMT", "yyyyMMdd"),
    "DTEND;VALUE=DATE:" + Utilities.formatDate(endDate, "GMT", "yyyyMMdd"),
    summary,
    description,
    options
  );
}

/**
 * Formats a date as a compact iCal UTC timestamp (e.g. 20240131T080000Z).
 *
 * @param {Date|string|number} date - The date to format.
 * @return {string} The iCal-formatted timestamp.
 */
function toICalDate_(date) {
  return Utilities.formatDate(new Date(date), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'").replace(
    /[-:]/g,
    ""
  );
}

// Source: src/roster/legend.js
function legendLabelTokens_(value) {
  return Array.from(
    String(value || "")
      .normalize("NFC")
      .replace(/([\p{L}])(\d)/gu, "$1 $2")
      .replace(/(\d)([\p{L}])/gu, "$1 $2")
      .matchAll(/[\p{L}\p{N}]+\.?/gu),
    function (match) {
      return { raw: match[0], key: match[0].replace(/\.$/, "").toUpperCase() };
    }
  );
}

/** A zero-length definition is not a duty. Missing legacy clocks stay date-only. */
function isTeamDutyDefinition_(entry) {
  if (!entry) return false;
  if (typeof entry.durationRawInMinutes == "number") return entry.durationRawInMinutes > 0;
  if (entry.isFlexiTime === true) return true;
  if (entry.from && entry.to && Number.isInteger(entry.followingDaysTo)) {
    var start = entry.from.hourOfDay * 60 + entry.from.minute;
    var end = entry.followingDaysTo * 1440 + entry.to.hourOfDay * 60 + entry.to.minute;
    return Number.isFinite(start) && Number.isFinite(end) && end > start;
  }
  return true;
}

/** Normalize the supplied metadata; no named shift/vehicle/role dictionary. */
function legendDefinition_(entry) {
  if (!isTeamDutyDefinition_(entry) || typeof entry.longName != "string") return null;
  var type = String(entry.shiftTypeName || "")
    .normalize("NFC")
    .toUpperCase();
  var tokens = legendLabelTokens_(entry.longName).filter(function (token) {
    return (
      /\d/.test(token.key) ||
      /^\p{L}$/u.test(token.key) ||
      token.key.length < 2 ||
      !type.startsWith(token.key)
    );
  });
  if (!tokens.length) return null;
  return {
    key: tokens
      .map(function (token) {
        return token.key;
      })
      .join(" "),
    tokens: tokens,
    type: type,
    clocks: JSON.stringify([
      entry.from && entry.from.hourOfDay,
      entry.from && entry.from.minute,
      entry.to && entry.to.hourOfDay,
      entry.to && entry.to.minute,
      entry.followingDaysTo,
      entry.isFlexiTime === true
    ])
  };
}

/** Learn only supported label variants; identifiers and unknown qualifiers are retained. */
function buildLegendAssignments_(rosters) {
  var records = [],
    bySignature = new Map(),
    catalog = new WeakMap();
  function identity(token) {
    return /\d/.test(token.key) || /^\p{L}$/u.test(token.key);
  }
  rosters.forEach(function (roster) {
    roster.legendDuties.entries.forEach(function (entry) {
      var definition = legendDefinition_(entry);
      if (!definition) return;
      var signature = JSON.stringify([definition.key, definition.type, definition.clocks]);
      var record = bySignature.get(signature);
      if (!record) {
        record = Object.assign(definition, { variants: [], entries: [] });
        bySignature.set(signature, record);
        records.push(record);
      }
      record.entries.push(entry);
    });
  });
  var patterns = new Map();
  records.forEach(function (longer) {
    checkSyncBudget_();
    records.forEach(function (shorter) {
      if (
        longer.tokens.length <= shorter.tokens.length ||
        longer.type != shorter.type ||
        longer.tokens[0].key != shorter.tokens[0].key
      )
        return;
      var position = 0,
        extras = [],
        lastIdentifier = -1;
      shorter.tokens.forEach(function (token, i) {
        if (identity(token)) lastIdentifier = i;
      });
      for (var token of longer.tokens) {
        if (position < shorter.tokens.length && token.key == shorter.tokens[position].key)
          position++;
        else {
          if (identity(token) || position <= lastIdentifier) return;
          extras.push(token);
        }
      }
      if (position != shorter.tokens.length || !extras.length) return;
      var pattern = extras
        .map(function (token) {
          return token.key;
        })
        .join(" ");
      var anchored = shorter.tokens.filter(identity).length >= 2;
      var sameClocks = longer.clocks == shorter.clocks;
      if (anchored && sameClocks) {
        var bases = patterns.get(pattern) || new Set();
        bases.add(shorter.key);
        patterns.set(pattern, bases);
      }
      longer.variants.push({
        base: shorter,
        extras: extras,
        pattern: pattern,
        anchored: anchored && sameClocks
      });
    });
  });
  records.sort(function (a, b) {
    return a.tokens.length - b.tokens.length;
  });
  records.forEach(function (record) {
    var supported = record.variants.filter(function (variant) {
      return (
        variant.anchored ||
        (patterns.has(variant.pattern) && patterns.get(variant.pattern).size >= 2)
      );
    });
    var longest = supported.reduce(function (n, variant) {
      return Math.max(n, variant.base.tokens.length);
    }, 0);
    supported = supported.filter(function (variant) {
      return variant.base.tokens.length == longest;
    });
    var roots = new Set(
      supported.map(function (variant) {
        return variant.base.assignment.key;
      })
    );
    var assignment = { key: record.key, qualifiers: [] };
    if (roots.size == 1) {
      var variant = supported[0];
      assignment = {
        key: variant.base.assignment.key,
        qualifiers: variant.base.assignment.qualifiers.concat(
          variant.extras
            .map(function (token) {
              return token.raw;
            })
            .join(" ")
        )
      };
    }
    record.assignment = assignment;
    record.entries.forEach(function (entry) {
      catalog.set(entry, assignment);
    });
  });
  return catalog;
}

/**
 * The UTC offset (ms) of a civil day in the given zone, or null when it changes during
 * that day (daylight saving). Google's date functions are slow in Apps Script; asking
 * once per day instead of four times per duty saves thousands of calls per team view.
 */
function civilDayOffset_(day, timeZone) {
  var midnight = function (date) {
    var key = timeZone + "|" + date;
    if (!zoneDayOffsetCache.has(key))
      zoneDayOffsetCache.set(
        key,
        Date.parse(date + "T00:00:00Z") -
          Utilities.parseDate(date, timeZone, "yyyy-MM-dd").getTime()
      );
    return zoneDayOffsetCache.get(key);
  };
  var next = new Date(Date.parse(day + "T00:00:00Z") + 86400000).toISOString().slice(0, 10);
  var offset = midnight(day);
  return offset === midnight(next) ? offset : null;
}

/** Scheduled wall-clock times from the legend; valuation/break minutes are not event bounds. */
function legendDutySpan_(entry, date) {
  var timeZone = rosterZone_();
  var key = JSON.stringify([
    date,
    timeZone,
    entry.from,
    entry.to,
    entry.followingDaysTo,
    entry.isFlexiTime
  ]);
  if (!legendDutySpanCache.has(key)) {
    function validClock(clock) {
      return (
        clock &&
        Number.isInteger(clock.hourOfDay) &&
        clock.hourOfDay >= 0 &&
        clock.hourOfDay <= 23 &&
        Number.isInteger(clock.minute) &&
        clock.minute >= 0 &&
        clock.minute <= 59
      );
    }
    function instant(clock, offset) {
      var civil = new Date(date + "T00:00:00Z");
      civil.setUTCDate(civil.getUTCDate() + offset);
      if (isNaN(civil.getTime())) return null;
      var wall =
        civil.toISOString().slice(0, 10) +
        "T" +
        String(clock.hourOfDay).padStart(2, "0") +
        ":" +
        String(clock.minute).padStart(2, "0") +
        ":00";
      // On days without a clock change every wall time exists exactly once.
      var offset = civilDayOffset_(wall.slice(0, 10), timeZone);
      if (offset !== null) return Date.parse(wall + "Z") - offset;
      var pattern = "yyyy-MM-dd'T'HH:mm:ss";
      var value = Utilities.parseDate(wall, timeZone, pattern).getTime();
      // A nonexistent local time or a repeated DST hour is not an unambiguous plan time.
      if (Utilities.formatDate(new Date(value), timeZone, pattern) != wall) return null;
      for (var delta of [-3600000, 3600000])
        if (Utilities.formatDate(new Date(value + delta), timeZone, pattern) == wall) return null;
      return value;
    }
    var span = null;
    if (
      entry.isFlexiTime !== true &&
      validClock(entry.from) &&
      validClock(entry.to) &&
      Number.isSafeInteger(entry.followingDaysTo) &&
      entry.followingDaysTo >= 0
    ) {
      try {
        var start = instant(entry.from, 0),
          end = instant(entry.to, entry.followingDaysTo);
        if (start != null && end != null && end > start) span = { start: start, end: end };
      } catch (_) {
        // Invalid/incomplete definitions keep their date-only representation.
      }
    }
    legendDutySpanCache.set(key, span);
  }
  var cached = legendDutySpanCache.get(key);
  return cached ? { start: new Date(cached.start), end: new Date(cached.end) } : null;
}

// Source: src/roster/crew.js
function teamDayKey_(value) {
  var instant = new Date(value);
  if (value == null || isNaN(instant.getTime())) return null;
  var zone = rosterZone_();
  return Utilities.parseDate(
    Utilities.formatDate(instant, zone, "yyyy-MM-dd"),
    zone,
    "yyyy-MM-dd"
  ).getTime();
}

/** Build once from the already fetched legend and team rows. */
function buildTeamDutyIndex_(rosters) {
  if (teamCrewCache.has(rosters)) return teamCrewCache.get(rosters);
  var started = Date.now();
  var index = {
    catalog: buildLegendAssignments_(rosters),
    duties: [],
    byEmployee: new Map(),
    byAssignment: new Map()
  };
  var seen = new Map(),
    dates = new Map();
  rosters.forEach(function (roster) {
    var legend = teamDutyLegend_(roster);
    roster.data.forEach(function (person) {
      checkSyncBudget_();
      var employee = person && person.item1;
      if (!employee || employee.id == null || typeof employee.name != "string") return;
      ((person.item2 && person.item2.data) || []).forEach(function (day) {
        checkSyncBudget_();
        if (!day || !day.day) return;
        if (!dates.has(day.day.date)) {
          var key = teamDayKey_(day.day.date);
          dates.set(
            day.day.date,
            key == null
              ? null
              : {
                  key: key,
                  date: Utilities.formatDate(new Date(key), rosterTimeZone, "yyyy-MM-dd")
                }
          );
        }
        var normalized = dates.get(day.day.date);
        if (!normalized) return;
        var dayKey = normalized.key,
          date = normalized.date;
        teamDutyCodes_(day).forEach(function (code) {
          var entry = legend.get(code),
            assignment = entry && index.catalog.get(entry);
          if (!assignment) return;
          var span = legendDutySpan_(entry, date);
          var id = String(employee.id);
          var signature = JSON.stringify([
            id,
            dayKey,
            code,
            assignment.key,
            span && span.start.getTime(),
            span && span.end.getTime(),
            String(entry.shiftTypeName || "").toUpperCase()
          ]);
          var duty = seen.get(signature);
          if (!duty) {
            duty = {
              id: id,
              name: formatName_(employee.name),
              day: dayKey,
              code: code,
              assignment: assignment.key,
              type: String(entry.shiftTypeName || "").toUpperCase(),
              span: span,
              entries: new Set(),
              qualifiers: new Set()
            };
            seen.set(signature, duty);
            index.duties.push(duty);
            if (!index.byEmployee.has(id)) index.byEmployee.set(id, []);
            index.byEmployee.get(id).push(duty);
            if (!index.byAssignment.has(assignment.key)) index.byAssignment.set(assignment.key, []);
            index.byAssignment.get(assignment.key).push(duty);
          }
          duty.entries.add(entry);
          assignment.qualifiers.forEach(function (qualifier) {
            duty.qualifiers.add(qualifier);
          });
        });
      });
    });
  });
  teamCrewCache.set(rosters, index);
  logDuration_("Team index built (" + index.duties.length + " duties)", started);
  return index;
}

function teamDutyLegend_(roster) {
  return new Map(
    roster.legendDuties.entries.map(function (entry) {
      return [entry.shortName, entry];
    })
  );
}

function teamDutyCodes_(day) {
  // Explicit absence codes retained from the user's sync rules.
  return String(day.shortName || "")
    .split(",")
    .map(function (code) {
      return code.trim();
    })
    .filter(function (code) {
      return code && !/^(K|UL|-)$/i.test(code);
    });
}

/** Prefer the actual team row/code, then unambiguous timing or an exact catalog label. */
function resolveTeamDuty_(index, employeeId, day, code, entry, span) {
  var candidates = (index.byEmployee.get(String(employeeId)) || []).filter(function (duty) {
    return duty.day == day;
  });
  var byCode = candidates.filter(function (duty) {
    return duty.code == code;
  });
  if (byCode.length) candidates = byCode;
  var exact = candidates.filter(function (duty) {
    return duty.entries.has(entry);
  });
  if (exact.length) candidates = exact;
  if (span) {
    var matching = candidates.filter(function (duty) {
      return duty.span && duty.span.start < span.end && duty.span.end > span.start;
    });
    if (matching.length) candidates = matching;
  }
  var assignments = new Set(
    candidates.map(function (duty) {
      return duty.assignment;
    })
  );
  if (assignments.size == 1)
    return Object.assign({}, candidates[0], { span: span || candidates[0].span });
  var definition = legendDefinition_(entry);
  if (definition && index.byAssignment.has(definition.key))
    return { assignment: definition.key, span: span, day: day, type: definition.type };
  return null;
}

/** Crew overlap and relief boundaries use instants, never the wording of a shift type. */
function teamDescriptionLines_(index, employeeId, day, code, entry, span) {
  if (!index || !index.byEmployee || day == null) return [];
  var current = resolveTeamDuty_(index, employeeId, day, code, entry, span);
  if (!current) return [];
  var duties = index.byAssignment.get(current.assignment) || [];
  function names(rows) {
    var people = new Map();
    rows.forEach(function (duty) {
      if (duty.id == employeeId) return;
      var person = people.get(duty.id);
      if (!person) {
        person = { name: duty.name, qualifiers: new Set() };
        people.set(duty.id, person);
      }
      duty.qualifiers.forEach(function (qualifier) {
        person.qualifiers.add(qualifier);
      });
    });
    return Array.from(people.values(), function (person) {
      var qualifiers = Array.from(person.qualifiers).sort();
      return person.name + (qualifiers.length ? " (" + qualifiers.join(" / ") + ")" : "");
    }).sort(function (a, b) {
      return a.localeCompare(b, "de");
    });
  }
  var lines = [];
  if (addTeamPartner) {
    var partners = names(
      duties.filter(function (duty) {
        return current.span && duty.span
          ? duty.span.start < current.span.end && duty.span.end > current.span.start
          : !current.span && !duty.span && duty.day == current.day && duty.type == current.type;
      })
    );
    if (partners.length) lines.push("Team: " + partners.join(", "));
  }
  if (addRelief && current.span) {
    var relief = names(
      duties.filter(function (duty) {
        return duty.span && duty.span.start.getTime() == current.span.end.getTime();
      })
    );
    if (relief.length) lines.push("Ablösung: " + relief.join(", "));
  }
  return lines;
}

// Source: src/roster/team.js
/**
 * Reformats a roster name from "Lastname,  Firstname" to "Firstname Lastname" so
 * a comma-separated list of names is unambiguous.
 */
function formatName_(raw) {
  var name = String(raw).replace(/\s+/g, " ").trim();
  var comma = name.indexOf(",");
  if (comma == -1) return name;
  var last = name.slice(0, comma).trim();
  var first = name.slice(comma + 1).trim();
  return (first + " " + last).trim();
}

var teamCacheSeconds = 600;

/**
 * Only what the app reads from a team view: names, staff IDs, days, duty codes and the
 * duty legend. Remarks, requests and other details are dropped before use and caching.
 * Malformed parts keep their shape, so the checks of the calendar sources still apply.
 */
function reduceTeamRoster_(roster) {
  var clock = function (value) {
    return value && typeof value == "object"
      ? { hourOfDay: value.hourOfDay, minute: value.minute }
      : value;
  };
  var day = function (entry) {
    if (!entry || typeof entry != "object") return entry;
    return {
      day: entry.day && typeof entry.day == "object" ? { date: entry.day.date } : entry.day,
      shortName: entry.shortName
    };
  };
  return {
    legendDuties: {
      entries: roster.legendDuties.entries.map(function (entry) {
        return {
          shortName: entry.shortName,
          longName: entry.longName,
          shiftTypeName: entry.shiftTypeName,
          durationRawInMinutes: entry.durationRawInMinutes,
          isFlexiTime: entry.isFlexiTime,
          from: clock(entry.from),
          to: clock(entry.to),
          followingDaysTo: entry.followingDaysTo
        };
      })
    },
    data: roster.data.map(function (row) {
      if (!row || typeof row != "object") return row;
      var person = row.item1,
        days = row.item2;
      return {
        item1: person && typeof person == "object" ? { id: person.id, name: person.name } : person,
        item2:
          days && typeof days == "object"
            ? { data: Array.isArray(days.data) ? days.data.map(day) : days.data }
            : days
      };
    })
  };
}

/** Team views are the same for everyone on a roster server; all accounts share them briefly. */
function sharedTeamCache_() {
  return typeof CacheService != "undefined" && typeof CacheService.getScriptCache == "function"
    ? CacheService.getScriptCache()
    : null;
}

/** Scheme and host are case-insensitive; the path may name a separate installation. */
function teamCacheServer_(url) {
  return String(url)
    .trim()
    .replace(/^[a-z][a-z0-9+.-]*:\/\/[^/?#]*/i, function (origin) {
      return origin.toLowerCase();
    });
}

function teamCacheKey_(groupId, from, to) {
  return (
    "team-roster:v1:" +
    sha256Hex_(JSON.stringify([teamCacheServer_(rosterUrl), String(groupId), from, to]))
  );
}

function readSharedTeamRoster_(cache, key) {
  try {
    var raw = cache.get(key);
    if (!raw) return null;
    var text =
      raw.charAt(0) == "z"
        ? Utilities.ungzip(
            Utilities.newBlob(Utilities.base64Decode(raw.slice(1)), "application/x-gzip")
          ).getDataAsString()
        : raw.slice(1);
    var roster = JSON.parse(text);
    return roster &&
      Array.isArray(roster.data) &&
      roster.legendDuties &&
      Array.isArray(roster.legendDuties.entries)
      ? roster
      : null;
  } catch (err) {
    return null; // A damaged entry is fetched again.
  }
}

function writeSharedTeamRoster_(cache, key, roster) {
  try {
    var text = JSON.stringify(roster),
      value = "j" + text;
    // A cache value may hold 100 kB; larger team views are compressed first.
    if (byteLength_(value) > 90000 && typeof Utilities.gzip == "function")
      value =
        "z" +
        Utilities.base64Encode(
          Utilities.gzip(Utilities.newBlob(text, "application/json")).getBytes()
        );
    if (byteLength_(value) <= 95000) cache.put(key, value, teamCacheSeconds);
  } catch (err) {
    log_("Could not cache a team view: " + ((err && err.message) || err));
  }
}

/** Own-calendar annotations can use partial team data. */
function getTeamDutyIndex_(monthParams, from, to) {
  return buildTeamDutyIndex_(loadTeamRosters_(monthParams, from, to, false));
}

/** Include the following morning for last-day night relief without importing it. */
function teamRosterRequest_(months, includeRelief) {
  var lastEnd = new Date(months[months.length - 1].end);
  var lastDay = Utilities.formatDate(lastEnd, rosterTimeZone, "yyyy-MM-dd");
  var monthParams = months.map(function (month) {
    return month.begin;
  });
  var to = Utilities.parseDate(lastDay, rosterTimeZone, "yyyy-MM-dd").toISOString();
  if (includeRelief) {
    to = new Date(lastEnd.getTime() + 1).toISOString();
    if (monthParams.indexOf(to) == -1) monthParams.push(to);
  }
  return {
    monthParams: monthParams,
    from: months[0].begin,
    to: to,
    firstDay: Utilities.formatDate(new Date(months[0].begin), rosterTimeZone, "yyyy-MM-dd"),
    lastDay: lastDay
  };
}

/** Fetch only team views accessible to the signed-in ROSTER account. Cache within one run. */
function loadTeamRosters_(monthParams, from, to, strict) {
  var key = JSON.stringify([
    rosterUrl,
    rosterUserId,
    rosterUserToken,
    rosterPlanningGroups,
    monthParams,
    from,
    to
  ]);
  if (teamRosterCache[key]) return teamRosterCache[key];
  var rosters = [],
    complete = true,
    // Only groups the roster server itself names for this account may come from the
    // shared cache; manually entered groups are always asked for with the own token.
    sharedCache = null;
  // Determine which planning groups to query.
  var groups =
    typeof rosterPlanningGroups !== "undefined" &&
    rosterPlanningGroups &&
    rosterPlanningGroups.length
      ? rosterPlanningGroups.slice()
      : [];
  if (!groups.length) {
    sharedCache = sharedTeamCache_();
    // Ask preload about every month: it only reports the groups planned for the
    // month it is given, so querying just one month misses groups the user works
    // in only in other months. Union the group ids across all months, deduped.
    // Annotation tolerates partial data; calendar sources require a complete response.
    var seen = {};
    var months = monthParams || [];
    rosterFetchAll_(
      months.map((month) => ({
        endpoint: "team-duty/preload",
        method: "POST",
        payload: JSON.stringify({
          employeeId: rosterUserId,
          planningGroup: null,
          filterEmployee: 0,
          month: month
        })
      })),
      strict
    ).forEach((body, i) => {
      try {
        if (body == null) {
          complete = false;
          return;
        }
        var pre = JSON.parse(body);
        var discovered = ((pre.teamDuties || {}).data || {}).data;
        if (!Array.isArray(discovered)) throw new Error("Unvollständige Teamübersicht");
        discovered.forEach((g) => {
          if (g.idPlanninggroup != null && !seen[g.idPlanninggroup]) {
            seen[g.idPlanninggroup] = true;
            groups.push(g.idPlanninggroup);
          }
        });
      } catch (e) {
        if (e.syncPaused) throw e;
        if (strict) throw new Error("Die Teamübersicht konnte nicht vollständig geladen werden.");
        complete = false;
        log_("Team-duty preload failed for month " + months[i] + ": " + (e.message || e));
      }
    });
  }

  log_("Team-duty planning groups: " + (groups.length ? groups.join(", ") : "(none)"));

  var writeCache = sharedTeamCache_();
  var byGroup = groups.map(function (groupId) {
    return sharedCache
      ? readSharedTeamRoster_(sharedCache, teamCacheKey_(groupId, from, to))
      : null;
  });
  var missing = groups.filter(function (groupId, i) {
    return !byGroup[i];
  });
  log_("Team-duty planning groups from the shared cache: " + (groups.length - missing.length));

  // Annotations can use partial results. Calendar sources must stop on any failure
  // so missing groups never look like deleted duties.
  var fetchStarted = Date.now();
  rosterFetchAll_(
    missing.map((groupId) => ({
      endpoint: "team-duty/roster/" + rosterUserId,
      method: "POST",
      payload: JSON.stringify({
        teamDuty: { idPlanninggroup: groupId, from: from, to: to },
        filterEmployee: 0
      })
    })),
    strict
  ).forEach((body, i) => {
    try {
      if (body == null) {
        complete = false;
        return;
      }
      var roster = JSON.parse(body).data;
      if (
        !roster ||
        !Array.isArray(roster.data) ||
        !roster.legendDuties ||
        !Array.isArray(roster.legendDuties.entries) ||
        !roster.legendDuties.entries.every(function (entry) {
          return entry && typeof entry.shortName == "string";
        })
      )
        throw new Error("Unvollständige Teamübersicht");
      var reduced = reduceTeamRoster_(roster);
      byGroup[groups.indexOf(missing[i])] = reduced;
      if (writeCache)
        writeSharedTeamRoster_(writeCache, teamCacheKey_(missing[i], from, to), reduced);
    } catch (e) {
      if (e.syncPaused) throw e;
      if (strict) throw new Error("Die Teamübersicht konnte nicht vollständig geladen werden.");
      complete = false;
      log_(
        "Team-duty roster fetch failed for planning group " + missing[i] + ": " + (e.message || e)
      );
    }
  });
  logDuration_(
    "Team-duty rosters fetched (" + missing.length + " of " + groups.length + ")",
    fetchStarted
  );
  byGroup.forEach(function (roster) {
    if (roster) rosters.push(roster);
  });

  if (strict && !complete)
    throw new Error("Die Teamübersicht konnte nicht vollständig geladen werden.");
  if (complete) teamRosterCache[key] = rosters;
  return rosters;
}

// Source: src/roster/colleagues.js
/** Source rows use [source, calendar name, color, calendar ID, optional person name]. */
function teamRosterEmployeeId_(source) {
  var match = /^TEAM_ROSTER:([1-9][0-9]*)$/.exec(String(source));
  return match ? match[1] : null;
}

/** Colleague views always use the normal four-month window, never the caller's contract. */
function currentTeamRosters_(request) {
  if (!hasRosterCredentials_(rosterUserId, rosterUserToken))
    throw new Error("Bitte zuerst deinen eigenen Dienstplan-Zugang hinterlegen.");
  request =
    request || teamRosterRequest_(JSON.parse(generateRosterPayload_(undefined, true)), false);
  return loadTeamRosters_(request.monthParams, request.from, request.to, true);
}

function teamRosterPeople_(rosters) {
  var people = new Map();
  rosters.forEach(function (roster) {
    roster.data.forEach(function (row) {
      var person = row && row.item1;
      if (!person || !/^[1-9][0-9]*$/.test(String(person.id)) || typeof person.name != "string")
        throw new Error("Die Personendaten der Teamübersicht sind unvollständig.");
      var name = formatName_(person.name);
      if (name) people.set(String(person.id), { id: String(person.id), name: name });
    });
  });
  return Array.from(people.values()).sort(function (a, b) {
    return a.name.localeCompare(b.name, "de");
  });
}

/** Combine team dates with scheduled legend times; personal notes/requests stay excluded. */
function getColleagueRosterICal_(employeeId, skipTokenRefresh) {
  if (!skipTokenRefresh && !teamRosterTokenReady) {
    refreshRosterToken_();
    teamRosterTokenReady = true;
  }
  var request = teamRosterRequest_(JSON.parse(generateRosterPayload_(undefined, true)), addRelief);
  var rosters = currentTeamRosters_(request);
  var teamIndex = addTeamPartner || addRelief ? buildTeamDutyIndex_(rosters) : {};
  var people = teamRosterPeople_(rosters);
  var person = people.find(function (item) {
    return item.id == employeeId;
  });
  if (!person)
    throw new Error(
      "Die ausgewählte Person ist in deiner Teamübersicht nicht mehr sichtbar. Bestehende Kalendertermine bleiben erhalten."
    );
  var personTitle =
    person.name +
    (people.filter(function (item) {
      return item.name == person.name;
    }).length > 1
      ? " (" + person.id + ")"
      : "");
  var events = new Map();
  rosters.forEach(function (roster) {
    checkSyncBudget_();
    var legend = teamDutyLegend_(roster);
    roster.data
      .filter(function (row) {
        return String(row.item1.id) == employeeId;
      })
      .forEach(function (row) {
        if (!row.item2 || !Array.isArray(row.item2.data))
          throw new Error("Die Dienste der ausgewählten Person sind unvollständig.");
        row.item2.data.forEach(function (day) {
          checkSyncBudget_();
          if (!day || typeof day != "object")
            throw new Error("Ein Eintrag der Teamübersicht ist unvollständig.");
          var instant = new Date(day.day && day.day.date);
          if (
            !day.day ||
            !day.day.date ||
            isNaN(instant.getTime()) ||
            (day.shortName != null && typeof day.shortName != "string")
          )
            throw new Error("Ein Eintrag der Teamübersicht ist unvollständig.");
          var date = Utilities.formatDate(instant, rosterTimeZone, "yyyy-MM-dd");
          if (date < request.firstDay || date > request.lastDay) return;
          teamDutyCodes_(day).forEach(function (code) {
            var entry = legend.get(code);
            if (rosterIgnoreList.indexOf(code) != -1) return;
            if (!isTeamDutyDefinition_(entry)) return;
            var type = String(entry.shiftTypeName || "");
            var uid = "team-roster-" + employeeId + "-" + date + "-" + encodeURIComponent(code);
            var span = legendDutySpan_(entry, date);
            var crew = teamDescriptionLines_(
              teamIndex,
              employeeId,
              teamDayKey_(instant),
              code,
              entry,
              span
            );
            var content;
            if (span) {
              content = shiftBlockToICal_({
                uid: uid,
                shortName: code,
                summary: personTitle + " · " + code,
                longName: entry.longName || code,
                shiftType: type,
                start: span.start,
                end: span.end,
                extraLines: [
                  "Planzeiten aus der Legende der Teamübersicht. Individuelle Abweichungen sind darin nicht enthalten."
                ].concat(crew),
                transparent: true,
                keepSeparate: true
              });
            } else {
              var start = new Date(date + "T00:00:00Z"),
                end = new Date(start.getTime() + 86400000);
              content = generateAllDayICalEntry_(
                uid,
                start,
                end,
                personTitle + " · " + code,
                [
                  entry.longName || code,
                  type,
                  entry.isFlexiTime === true
                    ? "Flexible Dienstzeit laut Legende; ganztägig dargestellt."
                    : "Keine eindeutigen vollständigen Planzeiten in der Legende; ganztägig dargestellt."
                ]
                  .concat(crew)
                  .join("\n"),
                { transparent: true, keepSeparate: true }
              );
            }
            events.set(uid, content);
          });
        });
      });
  });
  return (
    "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//DienstplanSync/Team//DE\n" +
    Array.from(events.values()).join("") +
    "END:VCALENDAR"
  );
}

// Source: src/roster/shifts.js
/**
 * Converts one roster day into shift block objects (one per merged duty block).
 * Entries that share the same shift code and touch or overlap in time are
 * merged right away (e.g. a duty split into contiguous segments). Rendering
 * happens later in shiftBlockToICal_ so blocks can additionally be consolidated
 * across days first (see consolidateShiftBlocks_).
 */
function parseShiftToCal_(shift, legend, teamIndex) {
  legend = legend || {};
  // Merge consecutive entries that share the same shift and touch or overlap
  // (the next one starts at or before the end of the previous) into a block.
  var mergeWindow = consolidationEnabled_() ? consolidateGapMillis_() : 0;
  var blocks = [];
  var entries = shift.data.rosterDetails.entries.slice().sort(function (a, b) {
    return new Date(a.from) - new Date(b.from);
  });
  entries.forEach((dienst) => {
    if (rosterIgnoreList.includes(dienst.shortName)) return;
    var start = new Date(dienst.from);
    var end = new Date(dienst.to);
    var prev = blocks[blocks.length - 1];
    if (
      prev &&
      prev.shortName == dienst.shortName &&
      start.getTime() <= prev.end.getTime() + mergeWindow
    ) {
      if (end.getTime() > prev.end.getTime()) prev.end = end;
    } else {
      blocks.push({
        shortName: dienst.shortName,
        nameWorkplace: dienst.nameWorkplace,
        nameRole: dienst.nameRole,
        nameShiftType: dienst.nameShiftType,
        remark: dienst.bemerkungDienstelement,
        start: start,
        end: end
      });
    }
  });

  var tagesbemerkung = shift.data.rosterDetails.tagesbemerkung; // day-level note, applies to every block

  var result = [];
  blocks.forEach((block) => {
    var leg = legend[block.shortName] || {};
    var shiftType = leg.shiftTypeName || block.nameShiftType || "";

    // Title: "<code> | <workplace> (<role>)" — shift type lives in the description.
    var title =
      block.shortName +
      " | " +
      block.nameWorkplace +
      (block.nameRole ? " (" + block.nameRole + ")" : "");

    // Extra description lines (rendered after the name/type/duration lines):
    // remarks, day note, team partners and relief crew.
    var extraLines = [];
    if (block.remark) extraLines.push(block.remark);
    if (tagesbemerkung) extraLines.push(tagesbemerkung);

    var dayKey = teamDayKey_((shift.data.key && shift.data.key.day) || block.start);
    extraLines = extraLines.concat(
      teamDescriptionLines_(
        teamIndex,
        rosterUserId,
        dayKey,
        block.shortName,
        Object.assign({ longName: block.nameWorkplace, shiftTypeName: shiftType }, leg),
        { start: block.start, end: block.end }
      )
    );

    // On-call (Rufbereitschaft) standby should not block availability.
    var transparent = oncallAsFree && shiftType == "Rufbereitschaft";

    result.push({
      shortName: block.shortName,
      summary: title,
      longName: leg.longName,
      shiftType: shiftType,
      extraLines: extraLines,
      start: block.start,
      end: block.end,
      transparent: transparent
    });
  });
  return { blocks: result };
}

/**
 * Consolidates shift blocks that share the same shift code (and transparency)
 * and touch or overlap in time (optionally within consolidateMaxGapMinutes)
 * into a single block — e.g. a follow-up shift that starts before the previous
 * one ends because of transfer/driving time, or a chain of back-to-back
 * on-call days. Works across day and month boundaries. Matching by shift code
 * (not full title) means a relocated continuation at another workplace still
 * merges; the first segment's title wins and a differing later title is kept
 * as an "Anschluss:" description line. The merged block spans from the first
 * start to the latest end; description lines are unioned and the duration is
 * recomputed when rendering. Disabled via the consolidateEvents setting.
 *
 * @param {Array.<Object>} blocks - Blocks from parseShiftToCal_, any order.
 * @return {Array.<Object>} Consolidated blocks in chronological order.
 */
function consolidateShiftBlocks_(blocks) {
  var sorted = blocks.slice().sort(function (a, b) {
    return a.start - b.start;
  });
  if (!consolidationEnabled_()) return sorted;

  var gapMillis = consolidateGapMillis_();
  var result = [];
  // Track the newest block per shift code so an interleaved different shift
  // (e.g. a day shift inside an on-call week) doesn't break the chain it sits in.
  var lastByKey = {};
  sorted.forEach(function (block) {
    var key = block.shortName + "|" + block.transparent;
    var prev = lastByKey[key];
    if (prev && block.start.getTime() <= prev.end.getTime() + gapMillis) {
      if (block.end.getTime() > prev.end.getTime()) prev.end = block.end;
      if (
        block.summary != prev.summary &&
        prev.extraLines.indexOf("Anschluss: " + block.summary) == -1
      )
        prev.extraLines.push("Anschluss: " + block.summary);
      block.extraLines.forEach(function (line) {
        if (prev.extraLines.indexOf(line) == -1) prev.extraLines.push(line);
      });
    } else {
      result.push(block);
      lastByKey[key] = block;
    }
  });
  return result;
}

/**
 * Renders a shift block into a VEVENT. The description is composed of the full
 * shift name, "<type> · <duration>" (recomputed here so consolidated blocks
 * show their real total span) and any collected remark/team/relief lines.
 */
function shiftBlockToICal_(block) {
  var start = toICalDate_(block.start);
  var end = toICalDate_(block.end);
  var descLines = [];
  if (block.longName) descLines.push(block.longName);
  var typeAndDuration = [block.shiftType, formatDuration_(block.start, block.end)]
    .filter(Boolean)
    .join(" · ");
  if (typeAndDuration) descLines.push(typeAndDuration);
  descLines = descLines.concat(block.extraLines);
  // A shift is identified by its code and start. The UID also contains the end, as it
  // always has in existing calendars, so the identity is remembered separately:
  // overtime moves the end and must change the existing event, not add a second one.
  var uid = block.uid || block.shortName + start + end;
  if (!block.uid) rosterShiftKeys.set(uid, block.shortName + start);
  return generateICalEntry_(uid, start, end, block.summary, descLines.join("\n"), {
    transparent: block.transparent,
    keepSeparate: !!block.keepSeparate
  });
}

// Source: src/roster/feed.js
function getRosterStartDate_() {
  return callWithBackoff_(function () {
    var jsonResponse = JSON.parse(rosterFetch_("app/initial-preload", "POST"));
    var date = jsonResponse.contracts.data[0].eingestellt.split(".");
    return new Date(Date.UTC(Number(date[2]), Number(date[1]) - 1, Number(date[0])));
  });
}

/**
 * Requests complete roster months, including the previous month so overnight
 * shifts and late corrections remain available after rollover. Calendar-month
 * arithmetic avoids setMonth's day-29/30/31 overflow. Bounds use roster-local
 * midnight, independently of the Apps Script timezone and execution time.
 *
 * @param {Date} [now] - Reference time (defaults to the current time).
 */
function generateRosterPayload_(now, bounded) {
  var timeZone = rosterZone_();
  var currentMonth = Utilities.formatDate(now || new Date(), timeZone, "yyyy-MM").split("-");
  var year = Number(currentMonth[0]);
  var month = Number(currentMonth[1]) - 1;
  var startMonth = new Date(Date.UTC(year, month - 1, 1));
  var endMonth = new Date(Date.UTC(year, month + 3, 1)); // exclusive, through current month + 2
  if (addRosterSinceStart && !bounded) {
    var contractStart = getRosterStartDate_();
    if (contractStart == null || isNaN(contractStart.getTime()))
      throw new Error("Could not determine roster contract start date");
    startMonth = new Date(Date.UTC(contractStart.getUTCFullYear(), contractStart.getUTCMonth(), 1));
  }

  return buildRosterMonthPayload_(startMonth, endMonth, timeZone);
}

function buildRosterMonthPayload_(startMonth, endMonth, timeZone) {
  function monthStart(date) {
    return Utilities.parseDate(date.toISOString().slice(0, 10), timeZone, "yyyy-MM-dd");
  }
  var months = [];
  while (startMonth < endMonth) {
    var nextMonth = new Date(
      Date.UTC(startMonth.getUTCFullYear(), startMonth.getUTCMonth() + 1, 1)
    );
    months.push({
      employeeId: Number(rosterUserId),
      begin: monthStart(startMonth).toISOString(),
      end: new Date(monthStart(nextMonth).getTime() - 1).toISOString(),
      rosterViewMode: 4
    });
    startMonth = nextMonth;
  }

  return JSON.stringify(months);
}

function getRosterICal_(skipTokenRefresh) {
  if (!skipTokenRefresh && !teamRosterTokenReady)
    callWithBackoff_(function () {
      refreshRosterToken_();
      teamRosterTokenReady = true;
      return;
    });
  return callWithBackoff_(function () {
    var payload = generateRosterPayload_();
    var fetchStarted = Date.now();
    var jsonResponse = JSON.parse(rosterFetch_("rosters/preload", "POST", payload));
    logDuration_("Own roster fetched", fetchStarted);
    var icsContent =
      "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//GoogleKalenderSync/EN\nMETHOD:REQUEST\nNAME:Roster\nX-WR-CALNAME:Roster\n";

    var seenAbsences = {}; // fehlzeiten are repeated in every month, so dedupe by approvalId
    var allShiftBlocks = []; // collected across all months so consolidation can span month boundaries

    // Optionally look up which colleagues share each vehicle (team partners).
    var teamIndex = {};
    if (addTeamPartner || addRelief) {
      try {
        var request = teamRosterRequest_(JSON.parse(payload), addRelief);
        teamIndex = getTeamDutyIndex_(request.monthParams, request.from, request.to);
      } catch (e) {
        if (e.syncPaused) throw e;
        log_("Could not build team-duty index: " + (e.message || e));
      }
    }

    jsonResponse.forEach((month) => {
      // Build a lookup of shift code -> {longName, shiftTypeName, ...} for this month.
      var legend = {};
      var legendData =
        month.rosterUrlaubEinsatzwunsch &&
        month.rosterUrlaubEinsatzwunsch.data &&
        month.rosterUrlaubEinsatzwunsch.data.legendDuties;
      if (legendData && legendData.entries)
        legendData.entries.forEach((e) => {
          legend[e.shortName] = e;
        });

      month.rosterDetails.forEach((shift) => {
        var data = parseShiftToCal_(shift, legend, teamIndex);
        allShiftBlocks = allShiftBlocks.concat(data.blocks);
      });
      if (addRosterRequests) {
        month.rosterUrlaubEinsatzwunsch.data.data[0].item2.data.forEach((day) => {
          if (day.einsatzwunsch) {
            var date = toICalDate_(day.day.date);
            var summary =
              "Einsatzwunsch: " +
              (function (ew) {
                switch (ew) {
                  case 1:
                    return "Frei";
                  case 2:
                    return "Beliebiger Dienst";
                  case 3:
                    return "Bestimmter Dienst";
                  case 4:
                    return "Dienst ausschließen";
                  default:
                    return "Frei";
                }
              })(day.einsatzwunschWunsch);
            summary += " " + day.shortName;

            icsContent += generateICalEntry_(
              day.einsatzwunschWunsch + "-" + day.shortName + "-" + date,
              date,
              date,
              summary,
              day.kommentar
            );
          }
        });
      }
      if (addAbsences && month.fehlzeiten && month.fehlzeiten.data) {
        // Absences come from fehlzeiten, not rosterDetails, so rosterIgnoreList
        // (which suppresses the duplicate UL day-blocks in rosterDetails) does
        // not apply here.
        month.fehlzeiten.data.data.forEach((abs) => {
          if (seenAbsences[abs.approvalId]) return;
          seenAbsences[abs.approvalId] = true;
          var startDate = new Date(abs.von.date);
          var endDate = new Date(abs.bis.date);
          endDate.setDate(endDate.getDate() + 1); // DTEND is exclusive for all-day events
          var summary = "Abwesenheit: " + abs.absenceName + (abs.status != 0 ? " (beantragt)" : "");
          icsContent += generateAllDayICalEntry_(
            "absence-" + abs.approvalId,
            startDate,
            endDate,
            summary,
            ""
          );
        });
      }
    });

    //------------------------ Emit (consolidated) shift events ------------------------
    consolidateShiftBlocks_(allShiftBlocks).forEach((block) => {
      icsContent += shiftBlockToICal_(block);
    });

    if (updateAvailable == "true") {
      var date = toICalDate_(new Date());
      icsContent += generateICalEntry_(
        date + "-updateNotification",
        date,
        date,
        "Dienstplan Update verfügbar",
        "Es ist ein Update für das Dienstplanprogramm verfügbar. Bitte den Updateanweisungen in der Gruppe folgen."
      );
    }

    icsContent += "END:VCALENDAR";
    return icsContent;
  });
}

// Source: src/reports/sources.js
function rosterReportPayload_(year, timeZone) {
  return buildRosterMonthPayload_(
    new Date(Date.UTC(year, -1, 1)),
    new Date(Date.UTC(year, 13, 1)),
    timeZone
  );
}

function fetchRosterReportHolidays_(year, region) {
  var cache = typeof CacheService != "undefined" ? CacheService.getScriptCache() : null,
    key = "roster-tax-holidays:v1:" + year + ":" + region;
  var raw = cache && cache.get(key),
    data;
  if (raw) {
    try {
      data = JSON.parse(raw);
    } catch (_) {
      raw = null;
    }
  }
  if (!raw) {
    var response = UrlFetchApp.fetch(
      "https://feiertage-api.de/api/?jahr=" +
        year +
        "&nur_land=" +
        (region == "DE" ? "NATIONAL" : region),
      { muteHttpExceptions: true }
    );
    if (response.getResponseCode() != 200)
      throw new Error(
        "Die Feiertags-API ist derzeit nicht erreichbar (HTTP " +
          response.getResponseCode() +
          "). Bitte später erneut versuchen."
      );
    try {
      data = JSON.parse(response.getContentText());
    } catch (_) {
      throw new Error("Die Feiertags-API hat keine gültigen Daten geliefert.");
    }
  }
  if (!data || typeof data != "object" || Array.isArray(data))
    throw new Error("Die Feiertags-API hat ein unbekanntes Format geliefert.");
  var entries = Object.keys(data).map(function (name) {
    var item = data[name],
      date = item && item.datum;
    if (
      typeof date != "string" ||
      !/^\d{4}-\d{2}-\d{2}$/.test(date) ||
      Number(date.slice(0, 4)) != year ||
      isNaN(Date.parse(date)) ||
      new Date(date + "T12:00:00Z").toISOString().slice(0, 10) != date
    )
      throw new Error("Die Feiertags-API hat ein ungültiges Datum geliefert.");
    return { name: name.slice(0, 120), date: date, note: String(item.hinweis || "").slice(0, 600) };
  });
  var dates = new Set(
    entries.map(function (entry) {
      return entry.date;
    })
  );
  if (
    entries.length < 9 ||
    entries.length > 40 ||
    !["01-01", "05-01", "10-03", "12-25", "12-26"].every(function (day) {
      return dates.has(year + "-" + day);
    })
  )
    throw new Error(
      "Die Feiertagsdaten sind unvollständig. Es wird keine geschätzte Feiertagsberechnung verwendet."
    );
  if (cache && !raw) cache.put(key, JSON.stringify(data), 21600);
  return {
    provider: "feiertage-api.de",
    dates: entries
      .filter(function (entry) {
        return !entry.note.trim();
      })
      .map(function (entry) {
        return entry.date;
      }),
    notes: entries.filter(function (entry) {
      return !!entry.note.trim();
    }),
    region: region
  };
}

/** Cache only the small time/vehicle dataset, not raw roster responses or remarks. */
function loadAnnualRosterData_(year, timeZone, forceReload) {
  var cache = annualCache_();
  var identity = JSON.stringify([
    year,
    timeZone,
    rosterUserId,
    rosterUrl,
    rosterIgnoreList.slice().sort()
  ]);
  var key = "annual-roster:v3:" + sha256Hex_(identity),
    raw = cache && !forceReload && cache.get(key),
    data;
  if (raw) {
    try {
      data = JSON.parse(raw);
      if (Array.isArray(data.entries) && Array.isArray(data.standby))
        return Object.assign(data, { cacheHit: true });
    } catch (_) {}
  }
  if (
    !callWithBackoff_(function () {
      refreshRosterToken_();
      return true;
    })
  )
    throw new Error("Der Dienstplan-Zugang konnte nicht erneuert werden.");
  var body = rosterReportPayload_(year, timeZone);
  var response = callWithBackoff_(function () {
    return JSON.parse(rosterFetch_("rosters/preload", "POST", body));
  });
  if (!response) throw new Error("Die Jahresdaten konnten nicht vollständig geladen werden.");
  data = collectAnnualRosterEntries_(response, year, timeZone, rosterIgnoreList);
  data.sourceLoadedAt = new Date().toISOString();
  data.cacheHit = false;
  raw = JSON.stringify(data);
  if (cache && byteLength_(raw) < 90000) {
    var keys = JSON.parse(userPrp.getProperty("annualCacheKeys") || "[]").filter(function (item) {
      return item != key;
    });
    keys.push(key);
    while (keys.length > 12) cache.remove(keys.shift());
    // Record the key first so disconnect can always remove private cached data.
    userPrp.setProperty("annualCacheKeys", JSON.stringify(keys));
    cache.put(key, raw, 900);
  }
  return data;
}

function clearAnnualRosterCache_() {
  var cache = annualCache_();
  if (cache) cache.removeAll(JSON.parse(userPrp.getProperty("annualCacheKeys") || "[]"));
}

/** The per-account cache of Apps Script; absent only in a reduced test environment. */
function annualCache_() {
  return typeof CacheService != "undefined" && typeof CacheService.getUserCache == "function"
    ? CacheService.getUserCache()
    : null;
}

// Source: src/reports/entries.js
function rosterReportUnion_(entries) {
  var result = [];
  entries
    .slice()
    .sort(function (a, b) {
      return a.start - b.start || a.end - b.end;
    })
    .forEach(function (entry) {
      var last = result[result.length - 1];
      if (last && entry.start <= last.end) {
        last.end = Math.max(last.end, entry.end);
        last.vehicles = Array.from(new Set(last.vehicles.concat(entry.vehicles)));
        last.codes = Array.from(new Set(last.codes.concat(entry.codes)));
      } else
        result.push({
          start: entry.start,
          end: entry.end,
          vehicles: entry.vehicles.slice(),
          codes: entry.codes.slice()
        });
    });
  return result;
}

function collectAnnualRosterEntries_(months, year, timeZone, ignoreCodes) {
  var from = Utilities.parseDate(year + "-01-01", timeZone, "yyyy-MM-dd").getTime();
  var until = Utilities.parseDate(year + 1 + "-01-01", timeZone, "yyyy-MM-dd").getTime();
  var entries = [],
    standby = [],
    seen = new Set(),
    invalid = 0,
    duplicates = 0,
    visited = 0;
  if (!Array.isArray(months) || !months.length)
    throw new Error("Die Dienstplan-Antwort enthält keine Monatsdaten.");
  months.forEach(function (month) {
    checkSyncBudget_();
    if (!month || !Array.isArray(month.rosterDetails))
      throw new Error("Die Monatsdaten des Dienstplans sind unvollständig.");
    var legend = {};
    var definitions =
      (((month.rosterUrlaubEinsatzwunsch || {}).data || {}).legendDuties || {}).entries || [];
    definitions.forEach(function (item) {
      legend[item.shortName] = item;
    });
    (month.rosterDetails || []).forEach(function (shift) {
      (((shift.data || {}).rosterDetails || {}).entries || []).forEach(function (entry) {
        if (++visited % 100 == 0) checkSyncBudget_();
        if (visited > 50000)
          throw new Error("Die Quelle enthält zu viele Diensteinträge für eine Jahresübersicht.");
        if (ignoreCodes.indexOf(entry.shortName) != -1) return;
        var code = String(entry.shortName || "Ohne Kürzel");
        // K marks sickness, not worked time, regardless of the shift metadata.
        if (code.trim().toUpperCase() == "K") return;
        if (typeof entry.from != "string" || typeof entry.to != "string") {
          invalid++;
          return;
        }
        var start = new Date(entry.from).getTime(),
          end = new Date(entry.to).getTime();
        if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start) {
          invalid++;
          return;
        }
        if (end <= from || start >= until) return;
        // ROSTER calls this nameWorkplace, but the value identifies a vehicle.
        var vehicle = String(entry.nameWorkplace || "").trim() || "Fahrzeug fehlt";
        var key = JSON.stringify([start, end, vehicle, code]);
        if (seen.has(key)) {
          duplicates++;
          return;
        }
        seen.add(key);
        var type = String((legend[code] || {}).shiftTypeName || entry.nameShiftType || "")
          .trim()
          .toLowerCase();
        if (/^(frei|dienstfrei|freier tag|urlaub|krank|krankheit|abwesenheit)$/.test(type)) return;
        (type == "rufbereitschaft" ? standby : entries).push({
          start: start,
          end: end,
          vehicles: [vehicle],
          codes: [code]
        });
      });
    });
  });
  if (entries.length > 10000)
    throw new Error("Zu viele Diensteinträge für eine Jahresübersicht. Bitte die Quelle prüfen.");
  return {
    entries: entries,
    standby: standby,
    invalid: invalid,
    duplicates: duplicates,
    receivedMonths: months.length
  };
}

// Source: src/reports/calculation.js
function calculateAnnualRosterSummary_(data, options) {
  var year = options.year,
    timeZone = options.timeZone,
    holidays = new Set((options.holidayDates || []).concat(options.extraDates || []));
  var from = Utilities.parseDate(year + "-01-01", timeZone, "yyyy-MM-dd").getTime();
  var until = Utilities.parseDate(year + 1 + "-01-01", timeZone, "yyyy-MM-dd").getTime();
  var entries = data.entries,
    standby = data.standby;
  var planned = 0,
    rows = [],
    days = new Map(),
    boundaries = new Map(),
    overlap = false,
    now = Date.now();
  var sorted = entries.slice().sort(function (a, b) {
    return a.start - b.start;
  });
  var furthest = -Infinity;
  sorted.forEach(function (entry) {
    if (entry.start < furthest) overlap = true;
    furthest = Math.max(furthest, entry.end);
  });
  function dateAt(instant) {
    return Utilities.formatDate(new Date(instant), timeZone, "yyyy-MM-dd");
  }
  function nextDay(day, offset) {
    return new Date(Date.parse(day + "T12:00:00Z") + offset * 86400000).toISOString().slice(0, 10);
  }
  function cuts(day) {
    if (!boundaries.has(day))
      boundaries.set(
        day,
        [0, 4, 6, 14, 20, 21]
          .map(function (hour) {
            return Utilities.parseDate(
              day + "T" + String(hour).padStart(2, "0") + ":00:00",
              timeZone,
              "yyyy-MM-dd'T'HH:mm:ss"
            ).getTime();
          })
          .concat(Utilities.parseDate(nextDay(day, 1), timeZone, "yyyy-MM-dd").getTime())
      );
    return boundaries.get(day);
  }
  function empty() {
    return {
      minutes: 0,
      nightMinutes: 0,
      night21Minutes: 0,
      night40Minutes: 0,
      sunday50Minutes: 0,
      holiday125Minutes: 0,
      holiday150Minutes: 0
    };
  }
  var metricKeys = Object.keys(empty());
  function add(to, values) {
    metricKeys.forEach(function (key) {
      to[key] += values[key];
    });
  }
  function special(day) {
    var md = day.slice(5);
    return md == "05-01" || md == "12-25" || md == "12-26" || md == "12-24";
  }
  var blocks = rosterReportUnion_(entries);
  blocks.forEach(function (block) {
    checkSyncBudget_();
    if (!options.includeFuture && block.end > now) {
      planned++;
      return;
    }
    var start = Math.max(block.start, from),
      end = Math.min(block.end, until),
      metrics = empty();
    if (end <= start) return;
    for (var day = dateAt(start); day <= dateAt(end - 1); day = nextDay(day, 1)) {
      var cut = cuts(day),
        previous = nextDay(day, -1),
        weekday = new Date(day + "T12:00:00Z").getUTCDay(),
        values = empty();
      for (var segment = 0; segment < cut.length - 1; segment++) {
        var duration =
          Math.max(0, Math.min(end, cut[segment + 1]) - Math.max(start, cut[segment])) / 60000;
        if (!duration) continue;
        values.minutes += duration;
        if (segment < 2 || segment >= 4) values.nightMinutes += duration;
        if (segment < 2 || segment == 5) values.night21Minutes += duration;
        var carry = segment == 0 && block.start < cut[0];
        if (carry) values.night40Minutes += duration;
        var md = day.slice(5),
          rate = 0;
        if ((special(day) && (md != "12-24" || segment >= 3)) || (carry && special(previous)))
          rate = 150;
        else if (
          holidays.has(day) ||
          (md == "12-31" && segment >= 3) ||
          (carry && (holidays.has(previous) || previous.slice(5) == "12-31"))
        )
          rate = 125;
        else if (weekday == 0 || (carry && weekday == 1)) rate = 50;
        if (rate == 150) values.holiday150Minutes += duration;
        else if (rate == 125) values.holiday125Minutes += duration;
        else if (rate == 50) values.sunday50Minutes += duration;
      }
      add(metrics, values);
      if (!days.has(day)) days.set(day, empty());
      add(days.get(day), values);
    }
    rows.push(
      Object.assign(
        {
          key: block.start + "-" + block.end,
          start: new Date(block.start).toISOString(),
          end: new Date(block.end).toISOString(),
          date: dateAt(start),
          startLocal: Utilities.formatDate(new Date(block.start), timeZone, "yyyy-MM-dd HH:mm"),
          endLocal: Utilities.formatDate(new Date(block.end), timeZone, "yyyy-MM-dd HH:mm"),
          fullMinutes: (block.end - block.start) / 60000,
          vehicles: block.vehicles.sort(),
          codes: block.codes.sort(),
          crossYear: block.start < from || block.end > until
        },
        metrics
      )
    );
  });
  var total = empty(),
    monthly = Array.from({ length: 12 }, function (_, i) {
      return Object.assign(
        { month: String(i + 1).padStart(2, "0"), days: 0, blocks: 0, over8: 0, over12: 0 },
        empty()
      );
    });
  days.forEach(function (values, day) {
    add(total, values);
    var month = monthly[Number(day.slice(5, 7)) - 1];
    add(month, values);
    month.days++;
  });
  rows.forEach(function (row) {
    var month = monthly[Number(row.date.slice(5, 7)) - 1];
    month.blocks++;
    if (row.fullMinutes > 480) month.over8++;
    if (row.fullMinutes > 720) month.over12++;
  });
  var standbyMinutes = rosterReportUnion_(standby).reduce(function (sum, entry) {
    return (
      sum +
      (!options.includeFuture && entry.end > now
        ? 0
        : Math.max(0, Math.min(entry.end, until) - Math.max(entry.start, from)) / 60000)
    );
  }, 0);
  return {
    year: year,
    timeZone: timeZone,
    region: options.region,
    includeFuture: options.includeFuture,
    generatedAt: new Date().toISOString(),
    requestedMonths: 14,
    receivedMonths: data.receivedMonths,
    rows: rows,
    monthly: monthly,
    holidayProvider: options.holidayProvider || "",
    holidayNotes: options.holidayNotes || [],
    sourceLoadedAt: data.sourceLoadedAt || new Date().toISOString(),
    cacheHit: !!data.cacheHit,
    totals: Object.assign(
      {
        blocks: rows.length,
        days: days.size,
        over8: rows.filter(function (row) {
          return row.fullMinutes > 480;
        }).length,
        over12: rows.filter(function (row) {
          return row.fullMinutes > 720;
        }).length,
        standbyMinutes: standbyMinutes
      },
      total
    ),
    warnings: {
      invalid: data.invalid,
      duplicates: data.duplicates,
      overlap: overlap,
      planned: planned
    }
  };
}

// Source: src/settings/preferences.js
// Program defaults can evolve without replacing any user's saved preferences.
var preferenceDefaults = {
  sourceCalendars: [["ROSTER", "Dienstplan", 5]],
  howFrequent: 15,
  onlyFutureEvents: false,
  addEventsToCalendar: true,
  modifyExistingEvents: true,
  removeEventsFromCalendar: true,
  removePastEventsFromCalendar: false,
  addCalToTitle: false,
  defaultAllDayReminder: -1,
  showSkipMessage: false,
  overrideVisibility: "",
  errorNotificationEmail: true,
  consolidateEvents: true,
  consolidateMaxGapMinutes: 0,
  ignoreEvents: [],
  addRosterToCal: true,
  addRosterSinceStart: false,
  rosterTimeZone: "Europe/Berlin",
  addRosterRequests: true,
  addAbsences: true,
  oncallAsFree: true,
  addTeamPartner: false,
  rosterPlanningGroups: [],
  addRelief: false,
  addYearSummary: false,
  summaryYear: "2024",
  rosterUrl: "https://dienstplan.drk-aachen.de:6100/api/",
  rosterIgnoreList: ["-", "UL"]
};

function encodePreferences_(settings) {
  return JSON.stringify(settings, function (key, value) {
    return value instanceof RegExp ? { $regex: value.source, $flags: value.flags } : value;
  });
}

function decodePreferences_(text) {
  return JSON.parse(text, function (key, value) {
    return value &&
      typeof value == "object" &&
      Object.keys(value).length == 2 &&
      typeof value.$regex == "string" &&
      typeof value.$flags == "string"
      ? new RegExp(value.$regex, value.$flags)
      : value;
  });
}

function readUserPreferences_(properties) {
  var saved = properties ? properties.preferences : userPrp.getProperty("preferences");
  var overrides = saved ? decodePreferences_(saved) : {};
  validatePreferences_(overrides);
  // Error emails are on by default since version 27. Settings saved before then stored
  // "off" only because it was preset, so it counts once the person saves it again.
  var choice = properties ? properties.errorEmailChoice : userPrp.getProperty("errorEmailChoice");
  if (overrides.errorNotificationEmail === false && !choice)
    delete overrides.errorNotificationEmail;
  return Object.assign(decodePreferences_(encodePreferences_(preferenceDefaults)), overrides);
}

function loadUserPreferences_() {
  var properties = userPrp.getProperties();
  var settings = readUserPreferences_(properties);
  Object.keys(settings).forEach(function (name) {
    globalThis[name] = settings[name];
  });
  rosterUserToken = properties.rosterUserToken;
  rosterUserId = properties.rosterUserId;
  syncOwnerId = currentUserEmail_();
  return settings;
}

/** Run from the owner's editor once to copy the existing personal installation. */
function migrateLegacySettings_() {
  requireAppOwner_();
  if (userPrp.getProperty("preferences"))
    throw new Error(
      "Vorhandene Einstellungen werden nicht überschrieben; die Migration ist bereits erfolgt"
    );
  var settings = {};
  Object.keys(preferenceDefaults).forEach(function (name) {
    settings[name] = globalThis[name] === undefined ? preferenceDefaults[name] : globalThis[name];
  });
  // Preserve the pre-v4 consolidation setting as well.
  if (typeof consolidateShifts !== "undefined" && typeof consolidateEvents === "undefined")
    settings.consolidateEvents = !!consolidateShifts;
  var values = { preferences: validatePreferences_(settings) };
  values.adoptLegacyEvents = "true";
  ["rosterUserToken", "rosterUserId"].forEach(function (name) {
    if (!userPrp.getProperty(name) && scriptPrp.getProperty(name))
      values[name] = scriptPrp.getProperty(name);
  });
  userPrp.setProperties(values, false);
  log_("Your settings and credentials were copied to your User Properties");
}

// Source: src/settings/validation.js
function validatePreferences_(settings) {
  if (!settings || typeof settings != "object" || Array.isArray(settings))
    throw new Error("Die Einstellungen haben ein ungültiges Format");
  Object.keys(settings).forEach(function (name) {
    if (!Object.prototype.hasOwnProperty.call(preferenceDefaults, name))
      throw new Error("Unbekannte Einstellung: " + name);
    var value = settings[name],
      fallback = preferenceDefaults[name];
    var valid = Array.isArray(fallback) ? Array.isArray(value) : typeof value == typeof fallback;
    if (name == "errorNotificationEmail")
      valid = typeof value == "boolean" || typeof value == "string";
    if (name == "summaryYear") valid = /^\d{4}$/.test(String(value));
    if (!valid || (typeof value == "number" && !Number.isFinite(value)))
      throw new Error("Ungültiger Wert für " + name);
  });
  if (
    settings.sourceCalendars &&
    !settings.sourceCalendars.every(function (row) {
      return (
        Array.isArray(row) &&
        row.length >= 2 &&
        row.length <= 5 &&
        typeof row[0] == "string" &&
        !!row[0].trim() &&
        typeof row[1] == "string" &&
        !!row[1].trim() &&
        (row[0] == "ROSTER" ||
          teamRosterEmployeeId_(row[0]) ||
          /^(https?|webcal):\/\//i.test(row[0])) &&
        (row.length < 4 || typeof row[3] == "string") &&
        (row.length < 5 ||
          (teamRosterEmployeeId_(row[0]) &&
            typeof row[4] == "string" &&
            !!row[4].trim() &&
            row[4].length <= 120))
      );
    })
  )
    throw new Error(
      "Jede Quelle benötigt einen Dienstplan oder eine Kalender-Adresse und einen Zielkalender"
    );
  if (
    settings.howFrequent !== undefined &&
    (settings.howFrequent < 1 || settings.howFrequent > 1440)
  )
    throw new Error("Das Intervall muss zwischen 1 und 1440 Minuten liegen");
  if (settings.consolidateMaxGapMinutes !== undefined && settings.consolidateMaxGapMinutes < 0)
    throw new Error("Die erlaubte Lücke darf nicht negativ sein");
  if (settings.rosterUrl !== undefined && !/^https:\/\//i.test(settings.rosterUrl))
    throw new Error("Die Dienstplan-Adresse muss mit https:// beginnen");
  if (settings.rosterTimeZone !== undefined) {
    try {
      Utilities.formatDate(new Date(), settings.rosterTimeZone, "yyyy-MM-dd");
    } catch (err) {
      throw new Error("Die Zeitzone ist ungültig");
    }
  }
  if (
    settings.rosterIgnoreList &&
    !settings.rosterIgnoreList.every(function (code) {
      return typeof code == "string";
    })
  )
    throw new Error("Dienstkürzel müssen als Text angegeben werden");
  if (
    settings.rosterPlanningGroups &&
    !settings.rosterPlanningGroups.every(function (group) {
      return (
        (typeof group == "number" && Number.isSafeInteger(group) && group >= 0) ||
        (typeof group == "string" && /^\d+$/.test(group))
      );
    })
  )
    throw new Error("Planungsgruppen müssen Nummern sein");
  if (
    settings.ignoreEvents &&
    !settings.ignoreEvents.every(function (rule) {
      if (typeof rule == "string" || rule instanceof RegExp) return true;
      if (!rule || typeof rule != "object" || Array.isArray(rule)) return false;
      if (
        Object.keys(rule).some(function (key) {
          return ["title", "weekday", "start", "end"].indexOf(key) == -1;
        })
      )
        return false;
      if (rule.title != null && typeof rule.title != "string" && !(rule.title instanceof RegExp))
        return false;
      if (
        rule.weekday != null &&
        !(
          (typeof rule.weekday == "number" &&
            Number.isInteger(rule.weekday) &&
            rule.weekday >= 1 &&
            rule.weekday <= 7) ||
          Object.prototype.hasOwnProperty.call(
            IGNORE_WEEKDAYS,
            String(rule.weekday).trim().toLowerCase()
          )
        )
      )
        return false;
      return ["start", "end"].every(function (key) {
        return rule[key] == null || /^([01]?\d|2[0-3]):[0-5]\d$/.test(String(rule[key]));
      });
    })
  )
    throw new Error("Eine Ausblendregel enthält ungültige Werte");
  var serialized = encodePreferences_(settings);
  if (byteLength_(serialized) > 9000)
    throw new Error("Die Einstellungen überschreiten den verfügbaren Speicherplatz");
  return serialized;
}

// Source: src/web/access.js
/** Run from the script editor. Private functions cannot be called by web clients. */
function setupSharedApp_() {
  var email = currentUserEmail_();
  if (!email) throw new Error("Bitte vor der Einrichtung bei Google anmelden");
  var owner = appOwnerEmail_();
  if (owner && owner != email) throw new Error("Nur der App-Inhaber kann die Einrichtung ändern");
  scriptPrp.setProperty("appOwnerEmail", email);
  if (!scriptPrp.getProperty("allowedUserEmails"))
    scriptPrp.setProperty("allowedUserEmails", JSON.stringify([email]));
  log_("Owner configured. Add your group's emails to allowedUserEmails in Script Properties");
}

function invitedEmails_() {
  return JSON.parse(scriptPrp.getProperty("allowedUserEmails") || "[]");
}

function requireAllowedUser_() {
  var owner = appOwnerEmail_();
  if (!owner)
    throw new Error("Der App-Inhaber muss zuerst setupSharedApp_ im Skripteditor ausführen");
  var email = currentUserEmail_();
  if (
    !email ||
    (email != owner &&
      !invitedEmails_().some(function (value) {
        return String(value).trim().toLowerCase() == email;
      }))
  )
    throw new Error("Dein Google-Konto ist für diese App nicht eingeladen");
  return email;
}

function requireAppOwner_() {
  var email = requireAllowedUser_();
  if (email != appOwnerEmail_()) throw new Error("Nur der App-Inhaber kann diese Aktion ausführen");
  return email;
}

function requireWebUser_() {
  var email = requireAllowedUser_();
  if (visitorEmail_() != email)
    throw new Error(
      "Die App muss als zugreifenden Nutzer ausgeführt werden; bitte den App-Inhaber kontaktieren"
    );
  return email;
}

function requireSyncUser_(event) {
  requireAllowedUser_();
  // One-shot triggers may disappear from the trigger list before their handler
  // runs. Accept only IDs explicitly issued for this account as well.
  if (
    event &&
    event.triggerUid &&
    ["initialSyncTriggerId", "syncContinuationTriggerId"].some(function (key) {
      return userPrp.getProperty(key) == event.triggerUid;
    })
  )
    return;
  if (
    event &&
    event.triggerUid &&
    ScriptApp.getProjectTriggers().some(function (trigger) {
      return (
        ["startSync", "continueSync_"].indexOf(trigger.getHandlerFunction()) != -1 &&
        trigger.getUniqueId() == event.triggerUid
      );
    })
  )
    return;
  requireWebUser_();
}

function getInvitations() {
  requireWebUser_();
  requireAppOwner_();
  return {
    owner: appOwnerEmail_(),
    emails: invitedEmails_(),
    health: readAccountHealth_(Date.now()),
    // Google runs other accounts' automatic runs only if they can read this project.
    projectUrl: "https://script.google.com/home/projects/" + ScriptApp.getScriptId() + "/edit"
  };
}

function saveInvitations(emails) {
  requireWebUser_();
  requireAppOwner_();
  if (!Array.isArray(emails)) throw new Error("Die E-Mail-Liste ist ungültig.");
  var normalized = Array.from(
    new Set(
      emails.map(function (value) {
        if (typeof value != "string" || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim()))
          throw new Error("Bitte gültige E-Mail-Adressen eingeben.");
        return value.trim().toLowerCase();
      })
    )
  ).sort();
  var text = JSON.stringify(normalized);
  if (byteLength_(text) > 9000) throw new Error("Die Zugriffsliste ist zu lang.");
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0))
    throw new Error("Die Zugriffsliste wird gerade bearbeitet. Bitte erneut versuchen.");
  try {
    scriptPrp.setProperty("allowedUserEmails", text);
  } finally {
    lock.releaseLock();
  }
  return getInvitations();
}

// Source: src/web/app.js
var publicSiteUrl = "https://justinbaumeyer.github.io/Google-Kalender-Sync/";

function escapeHtml_(value) {
  return String(value == null ? "" : value).replace(/[&<>"']/g, function (character) {
    return "&#" + character.charCodeAt(0) + ";";
  });
}

/** Explains which account was refused, since most refusals are a wrong Google profile. */
function accessDeniedPage_() {
  var email = "",
    configured = true;
  try {
    email = visitorEmail_();
    configured = !!appOwnerEmail_();
  } catch (err) {
    // The generic explanation below still applies.
  }
  var reason = !configured
    ? "Die App ist noch nicht eingerichtet. Der App-Inhaber muss zuerst die Einrichtung abschließen."
    : email
      ? "Das Konto <strong>" + escapeHtml_(email) + "</strong> ist für diese App nicht eingeladen."
      : "Bitte melde dich mit deinem eingeladenen Google-Konto an.";
  return (
    '<!doctype html><html lang="de"><head><base target="_top"><style>' +
    ":root{color-scheme:light dark;font:16px/1.55 system-ui,sans-serif}" +
    "body{margin:0;min-height:100vh;display:grid;place-items:center;background:#f3f6f3;color:#19392f;padding:1.25rem;box-sizing:border-box}" +
    "main{max-width:30rem;background:#fff;border:1px solid #dce4dd;border-radius:16px;padding:1.75rem}" +
    "h1{font-size:1.4rem;margin:0 0 .75rem}p{margin:.5rem 0}a{color:#146947}" +
    ".muted{color:#5b6e64;font-size:.9rem}" +
    "@media (prefers-color-scheme:dark){body{background:#0f1714;color:#e4ede7}main{background:#17221d;border-color:#2b3a33}a{color:#7fd0a6}.muted{color:#9db0a5}}" +
    "</style></head><body><main><h1>Dienstplan Sync</h1><p>Kein Zugriff. " +
    reason +
    '</p><p class="muted">Bist du im Browser mit mehreren Google-Konten angemeldet? Dann öffne den Link in einem Fenster, in dem nur das eingeladene Konto angemeldet ist, oder bitte den App-Inhaber, dieses Konto freizuschalten.</p>' +
    '<p class="muted"><a href="' +
    publicSiteUrl +
    '">Über Dienstplan Sync</a> · <a href="' +
    publicSiteUrl +
    'datenschutz.html">Datenschutz</a></p></main></body></html>'
  );
}

function doGet() {
  // Outside the try: when permissions are missing, Google ends this execution and
  // shows its consent screen, which must not be turned into the access notice below.
  requireFullAuthorization_();
  var output;
  try {
    requireWebUser_();
    repairStalledTriggers_();
    output = HtmlService.createHtmlOutputFromFile("Index");
  } catch (err) {
    output = HtmlService.createHtmlOutput(accessDeniedPage_());
  }
  output.setTitle("Dienstplan Sync");
  // Apps Script drops viewport tags from the HTML file; without this call phones
  // render the page at desktop width and the mobile layout never applies.
  output.addMetaTag("viewport", "width=device-width, initial-scale=1");
  try {
    output.setFaviconUrl(publicSiteUrl + "favicon.png");
  } catch (err) {
    // The icon is decoration only.
  }
  return output;
}

function getAppState() {
  var email = requireWebUser_();
  var saved = userPrp.getProperties();
  var settings = readUserPreferences_(saved);
  var enabled = hasTrigger_("startSync");
  var progress = saved.syncProgress ? JSON.parse(saved.syncProgress) : null;
  var isOwner = email == appOwnerEmail_();
  var availableVersion = Number(scriptPrp.getProperty("availableVersion")) || 0;
  return {
    version: version,
    email: email,
    settings: JSON.parse(encodePreferences_(settings)),
    rosterUserId: saved.rosterUserId || "",
    hasToken: !!saved.rosterUserToken,
    rosterConfigured: hasRosterCredentials_(saved.rosterUserId, saved.rosterUserToken),
    enabled: enabled,
    hasSavedPreferences: !!saved.preferences,
    isOwner: isOwner,
    nextRunAt: nextPlannedRun_(saved, enabled),
    lastStats: saved.lastSyncStats ? JSON.parse(saved.lastSyncStats) : null,
    connectionVerified:
      !!saved.rosterUserToken &&
      saved.connectionKey ==
        rosterConnectionKey_(saved.rosterUserToken, saved.rosterUserId, settings.rosterUrl),
    connectionCheckedAt: saved.connectionCheckedAt,
    credentialsRejected: saved.rosterAuthRejected == "true",
    triggersRepairedAt: saved.triggersRepairedAt || null,
    authorizationUrl: missingAuthorizationUrl_(),
    migrationImportedAt: saved.migrationImportedAt,
    needsCalendarSelection: saved.migrationNeedsCalendarSelection == "true",
    lastSync: saved.lastSyncAt,
    lastError: saved.lastSyncError,
    diagnostic: saved.lastSyncDiagnostic || "",
    pending: saved.syncPending == "true",
    migrationRemaining: Number(saved.migrationRemaining) || 0,
    progress: progress,
    history: readSyncHistory_(saved),
    stalled: syncLooksStalled_(saved, enabled, progress, settings.howFrequent),
    availableVersion: isOwner && availableVersion > version ? availableVersion : null,
    links: { site: publicSiteUrl, privacy: publicSiteUrl + "datenschutz.html" }
  };
}

function getSyncProgress() {
  requireWebUser_();
  var progress = userPrp.getProperty("syncProgress");
  return progress ? JSON.parse(progress) : null;
}

function saveAppSettings(input) {
  requireWebUser_();
  if (!input || typeof input != "object") throw new Error("Die Einstellungen fehlen");
  var preferences = decodePreferences_(JSON.stringify(input.settings));
  var serialized = validatePreferences_(preferences);
  // A saved choice about error emails is kept, even "off" (see readUserPreferences_).
  var values = { preferences: serialized, errorEmailChoice: "saved" };
  if (input.rosterUserId !== undefined) {
    if (input.rosterUserId !== "" && !/^\d+$/.test(String(input.rosterUserId)))
      throw new Error("Die Mitarbeiter-ID darf nur Ziffern enthalten");
    values.rosterUserId = String(input.rosterUserId);
  }
  if (input.rosterUserToken !== undefined) {
    if (typeof input.rosterUserToken != "string")
      throw new Error("Der Dienstplan-Token ist ungültig");
    var token = input.rosterUserToken.trim();
    if (byteLength_(token) > 9000)
      throw new Error("Der Dienstplan-Token ist zu lang zum Speichern");
    if (token) values.rosterUserToken = token;
  }
  withUserLock_("Eine Synchronisierung läuft. Bitte danach erneut speichern", function () {
    userPrp.setProperties(values, false);
    if (values.rosterUserToken) userPrp.deleteProperty("rosterAuthRejected");
    // The sync trigger interval follows saved preferences immediately.
    if (hasTrigger_("startSync")) install_();
  });
  return getAppState();
}

function stopSync() {
  requireWebUser_();
  deleteAllTriggers_(); // Apps Script returns only the current user's triggers.
  writeAccountHealth_({ enabled: false });
  userPrp.setProperty("syncPending", "false");
  publishSyncProgress_("Automatische Synchronisierung pausiert", undefined, undefined, "paused");
  userPrp.deleteProperty("initialSyncTriggerId");
  userPrp.deleteProperty("syncContinuationTriggerId");
  ["nextInitialAt", "nextContinuationAt", "nextRegularAt"].forEach(function (key) {
    userPrp.deleteProperty(key);
  });
  return getAppState();
}

function disconnectApp() {
  requireWebUser_();
  withUserLock_("Eine Synchronisierung läuft. Bitte danach erneut versuchen", function () {
    clearAnnualRosterCache_();
    userPrp.deleteAllProperties();
    deleteAllTriggers_();
    deleteAccountHealth_();
  });
  return getAppState();
}

// Source: src/web/drafts.js
function rosterConnectionKey_(token, employee, address) {
  return sha256Hex_(JSON.stringify([token || "", employee || "", address || ""]));
}

function markRosterConnectionVerified_() {
  userPrp.setProperties(
    {
      connectionKey: rosterConnectionKey_(rosterUserToken, rosterUserId, rosterUrl),
      connectionCheckedAt: new Date().toISOString()
    },
    false
  );
  userPrp.deleteProperty("rosterAuthRejected");
}

function nextPlannedRun_(saved, enabled) {
  var times = [saved.nextInitialAt, saved.nextContinuationAt].map(Number).filter(function (value) {
    return value > 0;
  });
  if (enabled && saved.nextRegularAt) {
    var next = Number(saved.nextRegularAt),
      interval = getValidTriggerFrequency_(readUserPreferences_(saved).howFrequent) * 60000;
    if (next < Date.now()) next += Math.ceil((Date.now() - next) / interval) * interval;
    times.push(next);
  }
  return times.length ? new Date(Math.min.apply(null, times)).toISOString() : null;
}

function withDraftPreferences_(input, operation) {
  requireWebUser_();
  return withUserLock_("Eine Synchronisierung läuft. Bitte danach erneut prüfen.", function () {
    var original = readUserPreferences_(),
      oldToken = rosterUserToken,
      oldId = rosterUserId,
      oldTimeZone = targetCalendarTimeZone;
    try {
      var draft = decodePreferences_(JSON.stringify((input && input.settings) || {}));
      validatePreferences_(draft);
      if (input && input.rosterUserToken !== undefined && typeof input.rosterUserToken != "string")
        throw new Error("Der Dienstplan-Token ist ungültig.");
      var settings = Object.assign({}, original, draft);
      validatePreferences_(settings);
      Object.keys(settings).forEach(function (name) {
        globalThis[name] = settings[name];
      });
      var savedId = userPrp.getProperty("rosterUserId"),
        savedToken = userPrp.getProperty("rosterUserToken");
      var draftToken = input && input.rosterUserToken && input.rosterUserToken.trim();
      rosterUserId =
        input && input.rosterUserId !== undefined ? String(input.rosterUserId).trim() : savedId;
      rosterUserToken = draftToken || savedToken;
      syncDeadline = Date.now() + 120000;
      resetRosterCaches_();
      // Diagnostics never rotate a draft token or write to a destination calendar.
      var mayRefresh =
        !draftToken &&
        savedToken &&
        rosterUserId == savedId &&
        settings.rosterUrl == original.rosterUrl;
      var refreshAttempted = false;
      return operation(settings, function (fetchOperation) {
        try {
          return fetchOperation ? fetchOperation() : getRosterICal_(true);
        } catch (err) {
          if (!mayRefresh || refreshAttempted || !err.rosterAuthRejected) throw err;
          refreshAttempted = true;
          if (!fetchOperation) return getRosterICal_();
          refreshRosterToken_();
          return fetchOperation();
        }
      });
    } catch (err) {
      if (err.syncPaused)
        throw new Error("Die Prüfung dauert zu lange. Bitte später erneut versuchen.");
      throw err;
    } finally {
      Object.keys(original).forEach(function (name) {
        globalThis[name] = original[name];
      });
      rosterUserToken = oldToken;
      rosterUserId = oldId;
      targetCalendarTimeZone = oldTimeZone;
      syncDeadline = 0;
      resetRosterCaches_();
    }
  });
}

function testRosterConnection(input) {
  return withDraftPreferences_(input, function (settings, fetchRoster) {
    if (!addRosterToCal) throw new Error("Bitte die Dienstplan-Quelle zum Prüfen aktivieren.");
    if (!hasRosterCredentials_(rosterUserId, rosterUserToken))
      throw new Error("Bitte Mitarbeiter-ID und Dienstplan-Token eingeben.");
    addRosterSinceStart = false;
    addTeamPartner = false;
    addRelief = false;
    var content = fetchRoster();
    if (!content) throw new Error("Der Dienstplan konnte nicht geladen werden.");
    var components = new ICAL.Component(ICAL.parse(content)).getAllSubcomponents("vevent");
    markRosterConnectionVerified_();
    return { ok: true, message: "Verbindung funktioniert.", eventsFound: components.length };
  });
}

function getTeamRosterPeople(input) {
  return withDraftPreferences_(input, function (settings, fetchRoster) {
    return fetchRoster(function () {
      return teamRosterPeople_(currentTeamRosters_());
    });
  });
}

// Source: src/web/preview.js
function previewEvents_(components, calendarName, calendarTimeZone) {
  var now = Date.now(),
    result = [],
    events = [],
    parents = new Map();
  components.forEach(function (component) {
    if (!component || typeof component.getFirstPropertyValue != "function") return;
    // Link exceptions by UID below instead of auto-linking every sibling exception.
    var event = new ICAL.Event(component, { strictExceptions: true, exceptions: [] });
    events.push(event);
    if (!event.isRecurrenceException()) parents.set(event.uid, event);
  });
  events.forEach(function (event) {
    if (event.isRecurrenceException() && parents.has(event.uid))
      parents.get(event.uid).relateException(event);
  });
  // All-day and floating times have no zone of their own; they mean the calendar's wall clock.
  function toDate(time) {
    var floating = time.zone && time.zone.tzid == "floating";
    return time.isDate || floating
      ? Utilities.parseDate(
          time.toString(),
          calendarTimeZone || rosterTimeZone,
          time.isDate ? "yyyy-MM-dd" : "yyyy-MM-dd'T'HH:mm:ss"
        )
      : time.toJSDate();
  }
  function instant(time) {
    return toDate(time).getTime();
  }
  function asString(time) {
    return time.isDate ? time.toString() : toDate(time).toISOString();
  }
  function add(event, start, end) {
    if (!start || !end || instant(end) < now) return;
    result.push({
      title: String(event.summary || "Ohne Titel"),
      start: asString(start),
      end: asString(end),
      allDay: !!start.isDate,
      calendar: calendarName
    });
  }
  events.forEach(function (event) {
    checkSyncBudget_();
    if (event.isRecurrenceException() && parents.has(event.uid)) return;
    if (event.isRecurring()) {
      var iterator = event.iterator(),
        occurrence,
        found = 0;
      for (var n = 0; n < 10000 && found < 5 && (occurrence = iterator.next()); n++) {
        if (n % 100 == 0) checkSyncBudget_();
        var details = event.getOccurrenceDetails(occurrence);
        if (instant(details.endDate) >= now) {
          add(details.item, details.startDate, details.endDate);
          found++;
        }
      }
    } else add(event, event.startDate, event.endDate);
  });
  return result
    .sort(function (a, b) {
      return a.start.localeCompare(b.start);
    })
    .slice(0, 5);
}

function previewSynchronization(input) {
  return withDraftPreferences_(input, function (settings, fetchRoster) {
    addTeamPartner = false;
    addRelief = false;
    var events = [],
      errors = [],
      cachedRoster;
    var calendars = condenseCalendarMap_(settings.sourceCalendars);
    calendars.forEach(function (calendar) {
      try {
        var sources = [];
        calendar[1].forEach(function (source) {
          if (source[0] == "ROSTER") {
            if (!isRosterSourceActive_()) return;
            if (cachedRoster === undefined) cachedRoster = fetchRoster();
            if (!cachedRoster) throw new Error("Der Dienstplan konnte nicht geladen werden.");
            sources.push([cachedRoster, source[1]]);
          } else {
            var fetched = fetchSourceCalendars_([source], fetchRoster);
            if (fetched.failedSources)
              throw new Error("Ein Kalender-Abo konnte nicht geladen werden.");
            sources = sources.concat(fetched.responses);
          }
        });
        ignoredEventIds = new Set();
        icsEventsIds = [];
        var timeZone = calendar[2]
          ? Calendar.CalendarList.get(calendar[2]).timeZone
          : Calendar.Settings.get("timezone").value;
        targetCalendarTimeZone = timeZone;
        events = events.concat(
          previewEvents_(consolidateVevents_(parseResponses_(sources)), calendar[0], timeZone)
        );
      } catch (err) {
        if (err.syncPaused) throw err;
        errors.push({ calendar: calendar[0], message: String(err.message || err).slice(0, 250) });
      }
    });
    return {
      events: events
        .sort(function (a, b) {
          return a.start.localeCompare(b.start);
        })
        .slice(0, 5),
      errors: errors,
      note: "Beispiele kommender Termine. Sehr lange Serien können in der Vorschau gekürzt sein. Es wurden keine Kalendertermine geschrieben."
    };
  });
}

// Source: src/web/annual.js
function getAnnualRosterSummary(input) {
  requireWebUser_();
  if (
    !input ||
    !isReportYear_(input.year) ||
    [
      "BW",
      "BY",
      "BE",
      "BB",
      "HB",
      "HH",
      "HE",
      "MV",
      "NI",
      "NW",
      "RP",
      "SL",
      "SN",
      "ST",
      "SH",
      "TH",
      "DE"
    ].indexOf(input.region) == -1 ||
    typeof input.includeFuture != "boolean"
  )
    throw new Error("Bitte ein Jahr ab 2020 und einen Feiertagskalender auswählen.");
  var dates = input.extraDates || [];
  if (input.forceReload !== undefined && typeof input.forceReload != "boolean")
    throw new Error("Die Aktualisierungsoption ist ungültig.");
  if (
    !Array.isArray(dates) ||
    dates.length > 100 ||
    !dates.every(function (day) {
      return (
        typeof day == "string" &&
        /^\d{4}-\d{2}-\d{2}$/.test(day) &&
        !isNaN(Date.parse(day)) &&
        new Date(day + "T12:00:00Z").toISOString().slice(0, 10) == day
      );
    })
  )
    throw new Error("Bitte gültige zusätzliche Feiertage auswählen.");
  return withDraftPreferences_({}, function () {
    if (!hasRosterCredentials_(rosterUserId, rosterUserToken))
      throw new Error("Bitte zuerst Mitarbeiter-ID und Dienstplan-Token speichern.");
    var reportTimeZone = "Europe/Berlin"; // German tax/tariff windows use local German time.
    var holidayData = fetchRosterReportHolidays_(input.year, input.region);
    var source = loadAnnualRosterData_(input.year, reportTimeZone, !!input.forceReload);
    var report = calculateAnnualRosterSummary_(source, {
      year: input.year,
      timeZone: reportTimeZone,
      region: input.region,
      extraDates: dates,
      includeFuture: input.includeFuture,
      holidayDates: holidayData.dates,
      holidayProvider: holidayData.provider,
      holidayNotes: holidayData.notes
    });
    report.stationSettings = readAnnualStationSettings_(input.year);
    return report;
  });
}

function saveAnnualStationSettings(input) {
  requireWebUser_();
  var settings = validateAnnualStationSettings_(input);
  return withUserLock_(
    "Eine Prüfung oder Synchronisierung läuft. Bitte danach speichern.",
    function () {
      userPrp.setProperty("annualStations:" + settings.year, JSON.stringify(settings));
      return Object.assign(settings, { saved: true });
    }
  );
}

function isReportYear_(year) {
  return Number.isInteger(year) && year >= 2020 && year <= 2099;
}

function validateAnnualStationSettings_(input) {
  var text = function (value, limit) {
    return typeof value == "string" && value.length <= limit && !/[\x00-\x1f]/.test(value);
  };
  if (
    !input ||
    !isReportYear_(input.year) ||
    !text(input.firstStation, 120) ||
    !Array.isArray(input.vehicles) ||
    input.vehicles.length > 100
  )
    throw new Error(
      "Bitte ein gültiges Jahr, die erste Tätigkeitsstätte und Fahrzeugzuordnungen angeben."
    );
  var seen = new Set();
  var vehicles = input.vehicles.map(function (row) {
    if (
      !Array.isArray(row) ||
      row.length != 2 ||
      !text(row[0], 240) ||
      !row[0].trim() ||
      !text(row[1], 120) ||
      seen.has(row[0].trim())
    )
      throw new Error("Die Fahrzeugzuordnung enthält ungültige oder doppelte Einträge.");
    seen.add(row[0].trim());
    return [row[0].trim(), row[1].trim()];
  });
  var result = { year: input.year, firstStation: input.firstStation.trim(), vehicles: vehicles };
  if (byteLength_(JSON.stringify(result)) > 8500)
    throw new Error("Die Wachenzuordnung ist zu groß. Bitte weniger eigene Zuordnungen speichern.");
  return result;
}

function readAnnualStationSettings_(year) {
  var raw = userPrp.getProperty("annualStations:" + year);
  if (raw) {
    try {
      var settings = validateAnnualStationSettings_(JSON.parse(raw));
      if (settings.year != year) throw new Error("Year mismatch");
      return Object.assign(settings, { saved: true });
    } catch (_) {
      return {
        year: year,
        firstStation: "",
        vehicles: [],
        saved: false,
        warning:
          "Die gespeicherte Wachenzuordnung konnte nicht gelesen werden. Bitte neu zuordnen und speichern."
      };
    }
  }
  return { year: year, firstStation: "", vehicles: [], saved: false };
}

// Source: src/web/migration.js
function importLegacyCopy(input) {
  var email = requireWebUser_();
  if (
    !input ||
    typeof input.code != "string" ||
    input.code.trim().length > 20000 ||
    input.code.trim().indexOf("DPSYNC1.") !== 0
  )
    throw new Error("Bitte einen gültigen Umzugscode aus deiner bisherigen Kopie einfügen.");
  var payload;
  try {
    payload = JSON.parse(
      Utilities.newBlob(Utilities.base64DecodeWebSafe(input.code.trim().slice(8))).getDataAsString(
        "UTF-8"
      )
    );
  } catch (err) {
    throw new Error("Der Umzugscode konnte nicht gelesen werden.");
  }
  if (
    !payload ||
    payload.format != "dienstplan-migration" ||
    payload.version !== 1 ||
    payload.email !== email
  )
    throw new Error("Der Umzugscode gehört nicht zu diesem Google-Konto.");
  if (
    !Number.isFinite(payload.expiresAt) ||
    payload.expiresAt < Date.now() ||
    payload.expiresAt > Date.now() + 86400000 + 60000
  )
    throw new Error("Der Umzugscode ist abgelaufen. Bitte in der bisherigen Kopie neu erstellen.");
  if (payload.oldSyncPaused !== true)
    throw new Error("Bitte zuerst die Synchronisierung in der bisherigen Kopie pausieren.");
  var settings = decodePreferences_(JSON.stringify(payload.settings));
  var serialized = validatePreferences_(settings);
  if (String(payload.rosterUserId || "") && !/^\d+$/.test(String(payload.rosterUserId)))
    throw new Error("Im Umzugscode fehlt eine gültige Mitarbeiter-ID.");
  if (typeof payload.rosterUserToken != "string" || byteLength_(payload.rosterUserToken) > 9000)
    throw new Error("Im Umzugscode ist der Dienstplan-Token ungültig.");
  withUserLock_("Eine Synchronisierung läuft. Bitte danach erneut importieren.", function () {
    if (userPrp.getProperty("preferences") && input.replaceExisting !== true)
      throw new Error(
        "Es sind bereits Einstellungen gespeichert. Bitte das Ersetzen ausdrücklich bestätigen."
      );
    var writable = new Set(
      readWritableCalendarRows_().map(function (cal) {
        return cal.id;
      })
    );
    settings.sourceCalendars.forEach(function (row) {
      if (row[3] && !writable.has(row[3]))
        throw new Error(
          "Ein bisheriger Zielkalender ist für dieses Konto nicht beschreibbar. Bitte Zugriff prüfen."
        );
    });
    deleteAllTriggers_();
    userPrp.setProperties(
      {
        preferences: serialized,
        rosterUserId: String(payload.rosterUserId || ""),
        rosterUserToken: payload.rosterUserToken,
        adoptLegacyEvents: "true",
        migrationNeedsCalendarSelection: String(
          settings.sourceCalendars.some(function (row) {
            return row.length < 4;
          })
        ),
        migrationSourceId: String(payload.sourceScriptId || ""),
        migrationImportedAt: new Date().toISOString(),
        syncPending: "false",
        migrationRemaining: "0"
      },
      false
    );
    [
      "syncResumeCalendar",
      "syncContinuationTriggerId",
      "initialSyncTriggerId",
      "syncProgress",
      "nextInitialAt",
      "nextContinuationAt",
      "nextRegularAt"
    ].forEach(function (key) {
      userPrp.deleteProperty(key);
    });
  });
  return getAppState();
}

/** The helper is embedded at build time and therefore matches the installed version. */
function getMigrationSnippet() {
  requireWebUser_();
  return migrationSnippetSource_;
}

// Source: migration/Export.gs (embedded as text)
var migrationSnippetSource_ = "/** Paste this entire file into a NEW file in the OLD private script copy.\n * Run prepareSharedAppMigration. No values or existing files need editing.\n * The generated code contains credentials and is never written to logs.\n */\nfunction prepareSharedAppMigration() {\n  var email = Session.getEffectiveUser().getEmail().trim().toLowerCase();\n  if (!email || Session.getActiveUser().getEmail().trim().toLowerCase() != email)\n    throw new Error(\"Bitte direkt im Skripteditor mit deinem Google-Konto ausführen.\");\n  var project = PropertiesService.getScriptProperties(), personal = PropertiesService.getUserProperties();\n  var owner = project.getProperty(\"appOwnerEmail\");\n  if (owner && owner != email) throw new Error(\"Bitte deine eigene bisherige Skriptkopie verwenden.\");\n  var scriptLock = LockService.getScriptLock(), userLock = LockService.getUserLock(), userLocked = false;\n  if (!scriptLock.tryLock(0)) throw new Error(\"Die bisherige Synchronisierung läuft noch. Bitte danach erneut versuchen.\");\n  try {\n    userLocked = userLock.tryLock(0);\n    if (!userLocked) throw new Error(\"Die bisherige Synchronisierung läuft noch. Bitte danach erneut versuchen.\");\n    var saved = personal.getProperties(), settings = saved.preferences ? JSON.parse(saved.preferences) : {};\n    var names = [\"sourceCalendars\",\"howFrequent\",\"onlyFutureEvents\",\"addEventsToCalendar\",\"modifyExistingEvents\",\"removeEventsFromCalendar\",\"removePastEventsFromCalendar\",\"addCalToTitle\",\"defaultAllDayReminder\",\"showSkipMessage\",\"overrideVisibility\",\"errorNotificationEmail\",\"consolidateEvents\",\"consolidateMaxGapMinutes\",\"ignoreEvents\",\"addRosterToCal\",\"addRosterSinceStart\",\"rosterTimeZone\",\"addRosterRequests\",\"addAbsences\",\"oncallAsFree\",\"addTeamPartner\",\"rosterPlanningGroups\",\"addRelief\",\"addYearSummary\",\"summaryYear\",\"rosterUrl\",\"rosterIgnoreList\"];\n    if (!saved.preferences) names.forEach(function(name) { if (globalThis[name] !== undefined) settings[name] = globalThis[name]; });\n    if (settings.consolidateEvents === undefined && typeof consolidateShifts !== \"undefined\") settings.consolidateEvents = !!consolidateShifts;\n    if (!Array.isArray(settings.sourceCalendars)) throw new Error(\"Die bisherigen Kalenderquellen konnten nicht gelesen werden.\");\n    var calendars = [], pageToken;\n    do {\n      var options = {showHidden:true,minAccessRole:\"writer\",maxResults:250};\n      if (pageToken) options.pageToken = pageToken;\n      var page = Calendar.CalendarList.list(options);\n      calendars = calendars.concat((page.items || []).filter(function(cal) { return [\"owner\",\"writer\"].indexOf(cal.accessRole) != -1; }));\n      pageToken = page.nextPageToken;\n    } while (pageToken);\n    settings.sourceCalendars = settings.sourceCalendars.map(function(row) {\n      var result = row.slice();\n      if (!result[3]) {\n        var matches = calendars.filter(function(cal) { return (cal.summaryOverride || cal.summary) == row[1]; });\n        if (matches.length == 1) { if (result.length < 3) result.push(null); result[3] = matches[0].id; }\n      }\n      return result;\n    });\n    var payload = {format:\"dienstplan-migration\",version:1,email:email,sourceScriptId:ScriptApp.getScriptId(),\n      expiresAt:Date.now()+86400000,oldSyncPaused:true,settings:settings,\n      rosterUserId:saved.rosterUserId || project.getProperty(\"rosterUserId\") || \"\",\n      rosterUserToken:saved.rosterUserToken || project.getProperty(\"rosterUserToken\") || \"\"};\n    var text = JSON.stringify(payload,function(key,value) {\n      return value instanceof RegExp ? {$regex:value.source,$flags:value.flags} : value;\n    });\n    var code = \"DPSYNC1.\" + Utilities.base64EncodeWebSafe(text,Utilities.Charset.UTF_8);\n    if (Utilities.newBlob(code).getBytes().length > 9000)\n      throw new Error(\"Die Einstellungen sind zu groß für einen Umzugscode. Bitte den App-Inhaber kontaktieren; es wurde nichts pausiert.\");\n    project.setProperty(\"DienstplanMigration\",\"Wird vorbereitet – noch nicht importieren\");\n    ScriptApp.getProjectTriggers().forEach(function(trigger) {\n      if ([\"startSync\",\"continueSync_\",\"install\",\"main\",\"checkForUpdates\"].indexOf(trigger.getHandlerFunction()) != -1)\n        ScriptApp.deleteTrigger(trigger);\n    });\n    project.setProperty(\"DienstplanMigration\",code);\n    Logger.log(\"Die bisherige Synchronisierung ist pausiert. Kopiere den Wert DienstplanMigration unter Projekteinstellungen → Skripteigenschaften in die gemeinsame App. Der Code enthält Zugangsdaten, ist 24 Stunden gültig und gehört nur in die gemeinsame App. Nach erfolgreichem Umzug die Eigenschaft und die Hilfsdatei löschen.\");\n  } finally {\n    if (userLocked) userLock.releaseLock();\n    scriptLock.releaseLock();\n  }\n}\n";
