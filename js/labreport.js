function init_rowevents() {
    let headertext = document.getElementsByTagName('H1')[0].innerText;
    if (!headertext.includes("Summary View")) {
        return;
    }
    var t = document.getElementById('labreport');
    var rows = t.rows; //rows collection - https://developer.mozilla.org/en-US/docs/Web/API/HTMLTableElement
    for (var i=0; i<rows.length; i++) {
        rows[i].onclick = function () {
            if (this.parentNode.nodeName == 'THEAD') {
                return;
            }
            let url = this.firstChild.innerText + ".html"
            window.location.href = url
        };
    }
}