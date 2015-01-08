

Get-ChildItem -Recurse -Include *.xls,*.xlsx | ForEach-Object {
    export_vba $_.Name


}

function export_vba($file_name){
    $file_path = Join-Path $PWD $file_name

    $excel = new-object -ComObject Excel.Application

    $excel.Workbooks.Open($file_path) | % {
        $_.VBProject.VBComponents | % {
         
            $file_path = create_export_file_name $_ $file_name
        }
    }

    $excel.Quit()
}

function create_export_directory($file_name) {
    $root_directory_name = 
    $directory = Join-Path $PWD $file_name
    

    $directories = @("module", "class", "form", "document_module")

    foreach($directories in $directory) {

    }

}


function create_export_file_name($vb_component, $export_target_file) {

    switch($vb_component.Type) 
    {
        1{ 
            $extension = ".bas"
            $directory = "module"
        }
        2{ 
            $extension = ".cls"
            $directory = "class"
        }
        3{
            $extension = ".frm"
            $directory = "form"
        }
        # TODO ActiveX Object
        11{
            $extension = ""
            $directory = "module"
        }
        100{
            $extension = ".cls"
            $directory = "dosument_module"
        }
    }
    $root_directory = [System.IO.Path]::GetFileNameWithoutExtension($export_target_file)

    $create_directory = Join-Path $PWD $root_directory
    $create_directory = Join-Path $create_directory $directory
    $file_name = $vb_component.Name + $extension
    $file_path = Join-Path $create_directory $file_name

    return $file_path
}