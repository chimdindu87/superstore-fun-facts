/* Stand-in for the Tableau Extensions API so the app runs as a plain web page (for Tableau Public).
   It answers the app's own queries from the baked snapshot in data.js, so app.js runs unchanged. */
(function () {
  'use strict';
  var D = window.FF_DATA || {};
  // Each query is recognised by the field aliases the app asks for (its first variant).
  var BY_ALIASES = {
    'oid|od|ctry|lines|sales': 'orders',
    'mode|days|ctry|lines|orders': 'ship',
    'rid|pn|sub|disc|sales|profit|qty': 'lines',
    'cid|nm|sales|profit|orders': 'cust',
    'products|customers|states': 'totals'
  };
  function rows(cols) {
    var keys = Object.keys(cols), n = keys.length ? cols[keys[0]].length : 0, out = new Array(n);
    for (var i = 0; i < n; i++) { var r = {}; for (var k = 0; k < keys.length; k++) { r[keys[k]] = cols[keys[k]][i]; } out[i] = r; }
    return out;
  }
  var source = {
    name: 'Sample - Superstore',
    queryAsync: function (q) {
      var key = (q && q.fields || []).map(function (f) { return f.fieldAlias; }).join('|');
      var name = BY_ALIASES[key];
      if (!name || !D[name]) { return Promise.reject(new Error('No snapshot for this query.')); }
      return Promise.resolve({ data: rows(D[name]) });
    }
  };
  window.tableau = {
    extensions: {
      initializeAsync: function () { return Promise.resolve(); },
      workbook: { getAllDataSourcesAsync: function () { return Promise.resolve([source]); } }
    }
  };
})();
