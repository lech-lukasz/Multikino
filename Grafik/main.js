setInterval(function () {
    var fetch = require('./FetchSchedule.js');
    fetch.fetchSchedule();
}, 1000 * 60 *30);
