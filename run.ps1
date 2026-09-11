$ErrorActionPreference = "Stop"

$key = Read-Host "Введите актуальный GigaChat Authorization Key" -AsSecureString
$ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($key)
try {
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    $env:GIGACHAT_CREDENTIALS = $plain
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
}

$env:GIGACHAT_SCOPE = if ($env:GIGACHAT_SCOPE) { $env:GIGACHAT_SCOPE } else { "GIGACHAT_API_PERS" }
$env:GIGACHAT_MODEL = if ($env:GIGACHAT_MODEL) { $env:GIGACHAT_MODEL } else { "GigaChat-2" }
$env:GIGACHAT_BASE_URL = if ($env:GIGACHAT_BASE_URL) { $env:GIGACHAT_BASE_URL } else { "https://api.giga.chat/v1" }
$env:GIGACHAT_VERIFY_SSL_CERTS = if ($env:GIGACHAT_VERIFY_SSL_CERTS) { $env:GIGACHAT_VERIFY_SSL_CERTS } else { "true" }
$env:GIGACHAT_TIMEOUT = if ($env:GIGACHAT_TIMEOUT) { $env:GIGACHAT_TIMEOUT } else { "180" }
$env:GIGACHAT_MAX_RETRIES = if ($env:GIGACHAT_MAX_RETRIES) { $env:GIGACHAT_MAX_RETRIES } else { "2" }

python test_gigachat.py
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

python main.py $args
