

Get-ChildItem -Recurse -Include *.xls,*.xlsx | ForEach-Object {
    export_vba $_.Name
}

function export_vba($file_name){
    create_export_directory $file_name

    $excel = new-object -ComObject Excel.Application
    $file_path = Join-Path $PWD $file_name

    $excel.Workbooks.Open($file_path) | % {
        $_.VBProject.VBComponents | % {
            $export_file_path = create_export_file_name $_ $file_name

            $_.Export($export_file_path)
        }
    }

    $excel.Quit()
}

function create_export_directory($file_name) {
    $extension_exclude_name =[System.IO.Path]::GetFileNameWithoutExtension($file_name)
    $file_path = Join-Path $PWD $extension_exclude_name

    # TODO ActiveX Object
    $sub_directories = @("module", "class", "form", "document_module")

    foreach($sub_directory in $sub_directories) {
        $create_directory_path = Join-Path $file_path $sub_directory

        create_directory $create_directory_path
    }

}

function create_directory($directory_path) {
    if(!(Test-Path $directory_path)){
        mkdir $directory_path
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
