# Bu PowerShell Script HTML dosyasından .com ile biten e-posta adreslerini çıkarır. 
# Belirtilen HTML dosyasını okuyarak içindeki .com ile biten e-posta adreslerini bulur ve bir txt dosyasına kaydeder.
    
# InputFile Analiz edilecek HTML dosyasının tam yoludur.
# OutputFile Bulunan e-postaların kaydedileceği dosya yoludur.
    
# Kullanımı
# .\HTML_Email_Extractor.ps1
    

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$InputFile = "C:\Temp\Yayınla _ Akış _ LinkedIn.html",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "C:\Temp\extracted_emails.txt"
)

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Remove-HTMLTags {
    param(
        [string]$HtmlText
    )
    
    Write-ColorOutput "HTML içeriği temizleniyor..." "Yellow"
    
    # HTML etiketlerini kaldır
    $cleanText = $HtmlText -replace '<script[^>]*>.*?</script>', '' # Script taglerini kaldır
    $cleanText = $cleanText -replace '<style[^>]*>.*?</style>', '' # Style taglerini kaldır
    $cleanText = $cleanText -replace '<[^>]+>', ' ' # Diğer HTML taglerini kaldır
    
    # HTML entity'lerini decode et
    $cleanText = $cleanText -replace '&amp;', '&'
    $cleanText = $cleanText -replace '&lt;', '<'
    $cleanText = $cleanText -replace '&gt;', '>'
    $cleanText = $cleanText -replace '&quot;', '"'
    $cleanText = $cleanText -replace '&#39;', "'"
    $cleanText = $cleanText -replace '&nbsp;', ' '
    $cleanText = $cleanText -replace '&#x27;', "'"
    $cleanText = $cleanText -replace '&#x2F;', '/'
    
    # Fazla boşlukları ve satır sonlarını temizle
    $cleanText = $cleanText -replace '\s+', ' '
    $cleanText = $cleanText.Trim()
    
    return $cleanText
}

function Get-ComEmails {
    param(
        [string]$Text
    )
    
    Write-ColorOutput ".com ile biten e-posta adresleri aranıyor..." "Yellow"
    
    # .com ile biten e-posta adresleri için regex pattern
    $emailPattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.com\b'
    
    # E-posta adreslerini bul
    $emails = [regex]::Matches($Text, $emailPattern) | ForEach-Object { $_.Value.ToLower() }
    
    # Tekrarları kaldır ve alfabetik olarak sırala
    $uniqueEmails = $emails | Sort-Object -Unique
    
    return $uniqueEmails
}

# Ana işlem başlangıcı
Write-ColorOutput "`n=== HTML E-posta Çıkarıcı (.com) ===" "Cyan"
Write-ColorOutput "Başlangıç zamanı: $(Get-Date)" "Gray"

# Input dosyasının varlığını kontrol et
if (-not (Test-Path $InputFile)) {
    Write-ColorOutput "Hata: Dosya bulunamadı: $InputFile" "Red"
    Write-ColorOutput "Lütfen dosya yolunun doğru olduğundan emin olun." "Yellow"
    exit 1
}

Write-ColorOutput "`nHTML dosyası okunuyor: $InputFile" "Green"

try {
    # HTML dosyasını UTF-8 encoding ile oku
    $htmlContent = Get-Content -Path $InputFile -Raw -Encoding UTF8 -ErrorAction Stop
    Write-ColorOutput "Dosya başarıyla okundu. Boyut: $($htmlContent.Length) karakter" "Green"
}
catch {
    Write-ColorOutput "Hata: HTML dosyası okunamadı - $_" "Red"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($htmlContent)) {
    Write-ColorOutput "Hata: HTML dosyası boş!" "Red"
    exit 1
}

# HTML içeriğini temizle
$cleanText = Remove-HTMLTags -HtmlText $htmlContent

# .com ile biten e-posta adreslerini çıkar
$foundEmails = Get-ComEmails -Text $cleanText

# Sonuçları göster
Write-ColorOutput "`n=== SONUÇLAR ===" "Cyan"

if ($foundEmails.Count -eq 0) {
    Write-ColorOutput "Hiç .com e-posta adresi bulunamadı." "Red"
}
else {
    Write-ColorOutput "Bulunan .com e-posta adresleri ($($foundEmails.Count) adet):" "Green"
    Write-ColorOutput "=" * 60 "Gray"
    
    $counter = 1
    foreach ($email in $foundEmails) {
        Write-ColorOutput "$counter. $email" "White"
        $counter++
    }
    
    # Sonuçları dosyaya kaydet
    try {
        # Output dizinini oluştur
        $outputDir = Split-Path $OutputFile -Parent
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        $outputContent = @()
        $outputContent += "HTML E-posta Çıkarma Raporu (.com)"
        $outputContent += "Kaynak Dosya: $InputFile"
        $outputContent += "Tarih: $(Get-Date)"
        $outputContent += "Bulunan .com e-posta sayısı: $($foundEmails.Count)"
        $outputContent += "=" * 60
        $outputContent += ""
        
        $counter = 1
        foreach ($email in $foundEmails) {
            $outputContent += "$counter. $email"
            $counter++
        }
        
        $outputContent += ""
        $outputContent += "--- Sadece E-posta Listesi ---"
        $outputContent += $foundEmails
        
        $outputContent | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
        Write-ColorOutput "`nSonuçlar '$OutputFile' dosyasına kaydedildi." "Green"
        
        # Sadece e-postaları içeren ayrı bir dosya oluştur
        $emailOnlyFile = $OutputFile -replace '\.txt$', '_emails_only.txt'
        $foundEmails | Out-File -FilePath $emailOnlyFile -Encoding UTF8 -Force
        Write-ColorOutput "Sadece e-postalar '$emailOnlyFile' dosyasına kaydedildi." "Green"
        
    }
    catch {
        Write-ColorOutput "Hata: Dosya kaydedilemedi - $_" "Red"
    }
}

# İstatistikler
Write-ColorOutput "`n=== İSTATİSTİKLER ===" "Cyan"
Write-ColorOutput "Kaynak dosya: $InputFile" "Gray"
Write-ColorOutput "Dosya boyutu: $([math]::Round($htmlContent.Length / 1024, 2)) KB" "Gray"
Write-ColorOutput "HTML içerik boyutu: $($htmlContent.Length) karakter" "Gray"
Write-ColorOutput "Temizlenmiş metin boyutu: $($cleanText.Length) karakter" "Gray"
Write-ColorOutput "Bulunan .com e-posta sayısı: $($foundEmails.Count)" "Gray"
Write-ColorOutput "Çıktı dosyası: $OutputFile" "Gray"
Write-ColorOutput "Bitiş zamanı: $(Get-Date)" "Gray"

# Pano'ya kopyalama seçeneği
if ($foundEmails.Count -gt 0) {
    $choice = Read-Host "`nBulunan e-postaları panoya kopyalamak ister misiniz? (E/H)"
    if ($choice -eq "E" -or $choice -eq "e") {
        try {
            $emailList = $foundEmails -join "`n"
            $emailList | Set-Clipboard
            Write-ColorOutput "E-postalar panoya kopyalandı!" "Green"
        }
        catch {
            Write-ColorOutput "Uyarı: Panoya kopyalanamadı - $_" "Yellow"
        }
    }
}

# Dosyayı açma seçeneği
if ($foundEmails.Count -gt 0) {
    $choice = Read-Host "`nÇıktı dosyasını açmak ister misiniz? (E/H)"
    if ($choice -eq "E" -or $choice -eq "e") {
        try {
            Start-Process $OutputFile
        }
        catch {
            Write-ColorOutput "Uyarı: Dosya açılamadı - $_" "Yellow"
        }
    }
}

Write-ColorOutput "`nScript başarıyla tamamlandı!" "Green"
