# Carregar o arquivo do Excel Dolar
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open("C:\Users\Bulbe\OneDrive - Bulbe\Power Automate\Conciliacao Bancaria\dolar.xlsx")

# Atualizar a planilha
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.UsedRange.Value2
$workbook.RefreshAll()

# Espera alguns segundos para que as tabelas sejam atualizadas
Start-Sleep -Seconds 10
$real = "R$"

# Criar as variáveis das células B7, B8 e B9 
$bid = $worksheet.Range("B9").Value2
$high = $worksheet.Range("B5").Value2
$low = $worksheet.Range("B6").Value2
$varBid = $worksheet.Range("B7").Value2
$pctChange = $worksheet.Range("B8").Value2
$ask = $worksheet.Range("B10").Value2
$create_date = $worksheet.Range("B12").Value2


# Fechar o arquivo e o Excel
$workbook.Close($true)
$excel.Quit()


#################################################################################################################################


# Carregar o arquivo do Excel Euro
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open("C:\Users\Bulbe\OneDrive - Bulbe\Power Automate\Conciliacao Bancaria\euro.xlsx")

# Atualizar a planilha
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.UsedRange.Value2
$workbook.RefreshAll()

# Espera alguns segundos para que as tabelas sejam atualizadas
Start-Sleep -Seconds 10
$real = "R$"
$pct = "%"

# Criar as variáveis das células B7, B8 e B9 
$bidEuro = $worksheet.Range("B9").Value2
$varBidEuro = $worksheet.Range("B7").Value2
$pctChangeEuro = $worksheet.Range("B8").Value2
$create_dateEuro = $worksheet.Range("B12").Value2

# Mostrar os resultados das variáveis criadas
Write-Host "bid: $bid"
Write-Host "high: $high"
Write-Host "low: $low"
Write-Host "varBid: $varBid"
Write-Host "pctChange: $pctChange"
Write-Host "ask: $ask"
Write-Host "create_date: $create_date"
Write-Host "bid Euro: $bidEuro"
Write-Host "varBid Euro: $varBidEuro"
Write-Host "pctChange Euro: $pctChangeEuro"
Write-Host "create_date Euro: $create_dateEuro"


# Fechar o arquivo e o Excel
$workbook.Close($true)
$excel.Quit()

###############################################################################################################################
$bidDolar = $real+$bid
$highDolar = $real+$high
$lowDolar = $real+$low
$varbidDolar = $real+$varBid
$askDolar = $real+$ask
$bidEuro2 = $real+$bidEuro
$varBidEuro2 = $real+$varBidEuro

$endpoint = "https://api.powerbi.com/beta/68745b6c-0c98-4490-a964-33265f634327/datasets/4e51f5c1-6f66-4631-9efd-025aaea92f92/rows?key=jEij%2Bk8Aijj2QoIVSZsQGQZI1%2BuFuX9FXTBpHBberZl7zY%2BwWfdH8gaTofaU6y448qXtZTiuC%2Bb%2B3EdhSGe6iw%3D%3D"
$payload = @{
"bid" = $bid
"high" = $high
"low" = $low
"varBid" = $varbid
"pctChange" = $pctChange
"ask" = $ask
"create_date" = $create_date
"bidEuro" = $bidEuro
"varBidEuro" = $varBidEuro
"pctChangeEuro" = $pctChangeEuro
"create_dateEuro" = $create_dateEuro
}
Invoke-RestMethod -Method Post -Uri "$endpoint" -Body (ConvertTo-Json @($payload))