param([string]$Py)

Add-Type -AssemblyName System.Windows.Forms

$d = New-Object System.Windows.Forms.OpenFileDialog
$d.Title  = '업데이트할 법규준수평가표를 선택하세요'
$d.Filter = 'Excel 파일 (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|모든 파일 (*.*)|*.*'

if ($d.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host '파일을 선택하지 않았습니다.'
    exit 0
}

& $Py main.py --update --file $d.FileName
exit $LASTEXITCODE
