function searchPhonebook() {
    var input, filter, table, rows, name, unit, ext, i;
    input = document.getElementById("searchInput");
    filter = input.value.toUpperCase();
    table = document.getElementById("phonebook");
    rows = table.getElementsByTagName("tr");
    for (i = 0; i < rows.length; i++) {
        name = rows[i].getElementsByTagName("td")[0];
        unit = rows[i].getElementsByTagName("td")[1];
        ext = rows[i].getElementsByTagName("td")[2];
        if (name || unit || ext) {
        if (name.innerHTML.toUpperCase().indexOf(filter) > -1 || unit.innerHTML.toUpperCase().indexOf(filter) > -1 || ext.innerHTML.toUpperCase().indexOf(filter) > -1) {
            rows[i].style.display = "";
        } else {
            rows[i].style.display= "none";
            }
        }
    }
    }