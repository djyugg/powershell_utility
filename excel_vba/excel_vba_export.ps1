

Get-ChildItem -Recurse -Include *.xls,*.xlsx | ForEach-Object {
    export_vba $_.Name


}

function export_vba($file_name){
    $file_path = Join-Path $PWD $file_name

    echo $file_path
    
    $excel = new-object -ComObject Excel.Application

    $excel.Workbooks.Open($file_path) | % {
        $_.VBProject.VBComponents | % {

        }
    }

}


function create_export_file_name($vb_component) {

    switch($vb_component.Type) 
    {
        1{ $extension = ".bas"}
        2{ $extension = ".cls"}
        3{ $extension = ".frm"}
        11 { $extension = "" }
        100{ $extension = ".bas"}
    }


    return [string]::Join($vb_component.Name, $extension)
}