var d = new Date();
var month = d.getMonth() + 1;
var day = d.getDate();
var year = d.getFullYear();

if (day == 1) {
    month = month - 1;
    day = 31;
}
else {
    day = day - 1;
}
var fulldate = month + "." + day + "." + year;
//set the date
$("#dateTo").val(fulldate);

//click the search button
$("#submit").click();

//setTimeout(function () {
    //$('#content').find('a')[2].click();
    //window.location.href = $('#content').find('a')[2].href;
//}, 2000);