Aşağıdaki PowerShell Script ile computer name'den, o bilgisayarı kullanan kişinin user name'ini, user OU'sunu ve computer OU'sunu bulabilirsiniz.

Scripti CheckComputer.ps1 olarak kaydebilirsiniz 

Kullanım için PowerShell'den CheckComputer.ps1 COMPUTER1923 gibi sorgulayabilirsiniz.

Son 30 günde bilgisayara giriş yapan kullanıcıları tespit eder,
Windows Event Log'dan Logon Type 2 (Interactive logon) kayıtlarını analiz eder,
Her kullanıcının hangi OU'da olduğunu gösterir.

# CheckComputer.ps1......................

param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName
)

# Yardım mesajı
if ($ComputerName -eq "/?" -or $ComputerName -eq "" -or $ComputerName -eq $null) {
    Write-Host "Kullanım: CheckComputer.ps1 ComputerName" -ForegroundColor Yellow
    Write-Host "Örnek: CheckComputer.ps1 COMPUTER01" -ForegroundColor Green
    exit
}

function CheckComputer {
    param([string]$CompName)
    
    try {
        Write-Host "===========================================" -ForegroundColor Cyan
        Write-Host "Bilgisayar Bilgileri: $CompName" -ForegroundColor Cyan
        Write-Host "===========================================" -ForegroundColor Cyan
        
        # Bilgisayarın bilgilerini al
        $computer = Get-ADComputer $CompName -Properties DistinguishedName, ManagedBy, Description
        
        if ($computer) {
            # Bilgisayarın OU'sunu göster
            $computerOU = $computer.DistinguishedName -replace "^CN=[^,]+,", ""
            Write-Host "`nBilgisayar OU'su:" -ForegroundColor Yellow
            Write-Host "  $computerOU" -ForegroundColor White
            
            # Bilgisayarın yöneticisini (ManagedBy) kontrol et
            if ($computer.ManagedBy) {
                try {
                    $manager = Get-ADUser $computer.ManagedBy -Properties DistinguishedName, Name, SamAccountName
                    $managerOU = $manager.DistinguishedName -replace "^CN=[^,]+,", ""
                    
                    Write-Host "`nBilgisayar Yöneticisi:" -ForegroundColor Yellow
                    Write-Host "  Kullanıcı: $($manager.Name) ($($manager.SamAccountName))" -ForegroundColor Green
                    Write-Host "  Kullanıcı OU'su: $managerOU" -ForegroundColor White
                }
                catch {
                    Write-Host "`nBilgisayar yöneticisi bilgisi alınamadı." -ForegroundColor Red
                }
            }
            else {
                Write-Host "`nBu bilgisayarda ManagedBy alanı boş." -ForegroundColor Orange
            }
            
            # Alternatif olarak bilgisayarı kullanan aktif kullanıcıları bul
            Write-Host "`nBilgisayarı Kullanan Aktif Kullanıcılar:" -ForegroundColor Yellow
            Write-Host "----------------------------------------" -ForegroundColor Gray
            
            # Son 30 günde giriş yapan kullanıcıları bul
            $thirtyDaysAgo = (Get-Date).AddDays(-30)
            $users = Get-WinEvent -ComputerName $CompName -FilterHashtable @{LogName='Security'; ID=4624; StartTime=$thirtyDaysAgo} -MaxEvents 100 -ErrorAction SilentlyContinue | 
                     Where-Object { $_.Message -match 'Logon Type:\s+2' } |
                     ForEach-Object { 
                         if ($_.Message -match 'Account Name:\s+([^\s]+)') {
                             $matches[1]
                         }
                     } | 
                     Where-Object { $_ -ne $env:COMPUTERNAME -and $_ -ne 'SYSTEM' -and $_ -ne 'ANONYMOUS LOGON' } |
                     Sort-Object -Unique
            
            if ($users) {
                foreach ($username in $users) {
                    try {
                        $user = Get-ADUser $username -Properties DistinguishedName -ErrorAction SilentlyContinue
                        if ($user) {
                            $userOU = $user.DistinguishedName -replace "^CN=[^,]+,", ""
                            Write-Host "  Kullanıcı: $username" -ForegroundColor Green
                            Write-Host "  OU: $userOU" -ForegroundColor White
                            Write-Host ""
                        }
                    }
                    catch {
                        # Kullanıcı bulunamazsa devam et
                    }
                }
            }
            else {
                Write-Host "  Event log'dan aktif kullanıcı bilgisi alınamadı." -ForegroundColor Orange
                Write-Host "  (WinRM servisi kapalı olabilir veya uzak erişim izni yoktur)" -ForegroundColor Orange
            }
            
            # Bilgisayar açıklaması varsa göster
            if ($computer.Description) {
                Write-Host "`nBilgisayar Açıklaması:" -ForegroundColor Yellow
                Write-Host "  $($computer.Description)" -ForegroundColor White
            }
        }
        else {
            Write-Host "`nBilgisayar bulunamadı!" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Hata: Bilgisayar '$CompName' bulunamadı veya erişim sorunu var." -ForegroundColor Red
        Write-Host "Detay: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Ana işlev çağrısı
CheckComputer $ComputerName

Write-Host "`n===========================================" -ForegroundColor Cyan
Write-Host "İşlem tamamlandı." -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
