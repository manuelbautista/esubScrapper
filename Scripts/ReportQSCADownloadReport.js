setTimeout(function () {
    $(".mix-hList_vertMiddle").click();
    //$(".tab-styledButton.left")[0].click();
    setTimeout(function () {
        var frame = document.getElementsByTagName("iframe")[0];
        frame.contentWindow.document.getElementsByTagName('a')[0].click();
        //$(".tab-styledButton.left")[0].click();
        //alert(frame.contentWindow.document);
    }, 2000); //2 Segundos
}, 2000); //2 Segundos
