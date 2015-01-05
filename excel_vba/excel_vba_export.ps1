

Get-ChildItem -Recurse -Include *.xls,*.xlsx | ForEach-Object {
    export_vba $_.Name


}

function export_vba($file_name){


}
