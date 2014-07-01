var alternate = 0
var standardbrowser = !document.all && !document.getElementById

if (standardbrowser)
    document.write('<form name="tick"><input type="text" name="tock" size="6"></form>')

function show() {
    if (!standardbrowser)
        var clockobj = document.getElementById ? document.getElementById("digitalclock") : document.all.digitalclock
    var Digital = new Date()
    var hours = Digital.getHours()
    var minutes = Digital.getMinutes()

    if (hours == 0) hours = 12
    if (hours.toString().length == 1)
        hours = "0" + hours
    if (minutes <= 9)
        minutes = "0" + minutes

    if (standardbrowser) {
        if (alternate == 0)
            document.tick.tock.value = hours + " : " + minutes + " "
        else
            document.tick.tock.value = hours + " " + minutes + " "
    }
    else {
        if (alternate == 0)
            clockobj.innerHTML = hours + "<font color='blue'>&nbsp;:&nbsp;</font>" + minutes + " "
        else
            clockobj.innerHTML = hours + "<font color='white'>&nbsp;:&nbsp;</font>" + minutes + " "
    }
    alternate = (alternate == 0) ? 1 : 0
    setTimeout("show()", 1000)
}
window.onload = show

function GetDate()
{
    var Mydate = new Date();
    var month = Mydate.getMonth() + 1;
    return Mydate.getDate() + "." + month + "." + Mydate.getYear();
}
