# OneDrive Entegrasyonu — Azure Kurulum

Telefonda login gerekmez. Kimlik bilgileri Vercel'de saklanır, yükleme otomatik yapılır.

---

## Adım 1: Azure Portal'da Uygulama Kaydet

1. https://portal.azure.com → Microsoft 365 hesabınla giriş yap
2. **"App registrations"** → **"New registration"**
3. Doldur:
   - **Name**: `WINDOFORM Uretim Scanner`
   - **Supported account types**: `Accounts in this organizational directory only`
   - Redirect URI: boş bırak (gerekmiyor)
4. **Register** tıkla

## Adım 2: Client ID ve Tenant ID'yi Not Al

- **Application (client) ID** → `AZURE_CLIENT_ID`
- **Directory (tenant) ID** → `AZURE_TENANT_ID`

## Adım 3: Client Secret Oluştur

1. Sol menü → **"Certificates & secrets"** → **"New client secret"**
2. Açıklama: `windoform-scanner`, Süre: 24 ay
3. **Add** → Değeri hemen kopyala (bir daha göremezsin!) → `AZURE_CLIENT_SECRET`

## Adım 4: API İzinleri (Application Permissions)

1. Sol menü → **"API permissions"** → **"Add a permission"**
2. **Microsoft Graph** → **Application permissions** (Delegated DEĞİL!)
3. Şunu ekle: `Files.ReadWrite.All`
4. **"Grant admin consent for [şirket adı]"** tıkla ✅

## Adım 5: Hedef Kullanıcının E-postası

Excel dosyalarının hangi kişinin OneDrive'ına gideceğini belirle
(örn: ofis sorumlusu, muhasebe vs.)
→ `ONEDRIVE_USER_EMAIL`

## Adım 6: Vercel'e Env Variables Ekle

Vercel Dashboard → Projen → Settings → Environment Variables:

| Key | Value |
|-----|-------|
| `AZURE_TENANT_ID` | Adım 2'den |
| `AZURE_CLIENT_ID` | Adım 2'den |
| `AZURE_CLIENT_SECRET` | Adım 3'ten |
| `ONEDRIVE_USER_EMAIL` | örn: `ahmet@windoform.com` |

Ardından Vercel'de **Redeploy** yap.

---

## Nasıl Çalışır?

```
Telefon → "OneDrive'a Kaydet" butonu
    ↓
Vercel sunucusu (gizli kimlik bilgileriyle token alır)
    ↓
Microsoft Graph API
    ↓
OneDrive/WINDOFORM/Uretim/
    ├── Uretim_11-06-2026_14-30.xlsx
    └── Uretim_11-06-2026_14-30_foto.jpg
```

Telefonda hiçbir login veya Microsoft hesabı gerekmez.
