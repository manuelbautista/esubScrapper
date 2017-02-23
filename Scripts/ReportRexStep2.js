setTimeout(function () {
    // $('*[onclick="loginToRR();return false;"]').click();
    var aLinks = document.getElementsByTagName('a');
    for (var i = 0; i < aLinks.length; i++) {
        if (aLinks[i].innerText == "Log in To Pay Per Call") {
            document.getElementsByTagName('a')[i].click();
        }
    }
    
}, 1000);