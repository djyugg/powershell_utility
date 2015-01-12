

Get-ChildItem -Recurse -Include *.xls,*.xlsx | ForEach-Object {
    export_vba $_.Name
}

function export_vba($file_name){
    $root_directory_name = [System.IO.Path]::GetFileNameWithoutExtension($file_name)
    $root_directory_path = Join-Path $PWD $root_directory_name

    create_export_directory $root_directory_path

    $excel = new-object -ComObject Excel.Application
    $file_path = Join-Path $PWD $file_name

    $excel.Workbooks.Open($file_path) | % {
        $_.VBProject.VBComponents | % {
            $export_file_path = create_export_file_name $_ $root_directory_path

            $_.Export($export_file_path)
        }
    }

    $excel.Quit()
}

function create_export_directory($root_directory_path) {
    # TODO ActiveX Object
    $sub_directories = @("module", "class", "form", "document_module")

    foreach($sub_directory in $sub_directories) {
        $create_directory_path = Join-Path $root_directory_path $sub_directory

        create_directory $create_directory_path
    }

}

function create_directory($directory_path) {
    if(!(Test-Path $directory_path)){
        mkdir $directory_path
    }
}

function create_export_file_name($vb_component, $root_directory_path) {

    switch($vb_component.Type)
    {
        1{
            $extension = ".bas"
            $sub_directory = "module"
        }
        2{
            $extension = ".cls"
            $sub_directory = "class"
        }
        3{
            $extension = ".frm"
            $sub_directory = "form"
        }
        # TODO ActiveX Object
        11{
            $extension = ""
            $sub_directory = "module"
        }
        100{
            $extension = ".cls"
            $sub_directory = "document_module"
        }
    }

    $create_directory = Join-Path $root_directory_path $sub_directory
    $export_file_name = $vb_component.Name + $extension
    $export_file_path = Join-Path $create_directory $export_file_name

    return $export_file_path
}
