var d = new Date();
var month = d.getMonth() + 1;
var day = d.getDate();
var year = d.getFullYear();

if (day == 1) {
    dateObj = new Date();
    dateObj.setDate(dateObj.getDate() - 1);
    //set previous date
    month = dateObj.getMonth() + 1;
    day = dateObj.getDate();
    year = d.getFullYear();
    //
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