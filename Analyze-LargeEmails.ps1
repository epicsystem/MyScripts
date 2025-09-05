# Son 7 günün 4MB üzeri maillerini analiz etme
$startDate = (Get-Date).AddDays(-7)
$endDate = Get-Date

# 4MB üzeri mailleri bulma
$largeMessages = Get-MessageTrace -StartDate $startDate -EndDate $endDate -PageSize 5000 | 
    Where-Object { $_.Size -gt 4194304 }  # 4MB = 4194304 bytes

# Rapor oluşturma
$report = @()
foreach ($message in $largeMessages) {
    $messageDetails = Get-MessageTraceDetail -MessageTraceId $message.MessageTraceId
    
    $report += [PSCustomObject]@{
        'Tarih' = $message.Received
        'Gönderen' = $message.SenderAddress
        'Alıcı' = $message.RecipientAddress
        'Konu' = $message.Subject
        'Boyut (MB)' = [math]::Round($message.Size/1MB, 2)
        'Durum' = $message.Status
    }
}

# Excel'e aktarma
$report | Export-Csv -Path "C:\Reports\LargeAttachments_4MB_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation -Encoding UTF8
