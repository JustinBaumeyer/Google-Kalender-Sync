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
