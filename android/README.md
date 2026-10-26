# Dükkan Asistanı Android

Bu Android uygulaması, depodaki Flask sitesini WebView içinde açar; site arayüzü, PIN girişi ve mevcut sunucu/API iş akışı aynen kullanılır. Uygulama ayrı bir veritabanı veya yeni backend gerektirmez.

## İlk kurulum

1. `android` klasörünü Android Studio ile açın ve Gradle eşitlemesinin tamamlanmasını bekleyin.
2. Android Studio'da Run'a basın veya APK oluşturun.
3. Uygulamanın ilk açılışında ngrok'un verdiği HTTPS adresini girin; örnek: `https://ornek.ngrok-free.app`.
4. Site açılınca mevcut personel PIN'inizle giriş yapın. Kamera barkod taraması istediğinde kamera iznini onaylayın.

## Kullanım

Telefon ve sunucunun çalıştığı bilgisayar internete bağlı olmalıdır. Bilgisayarda Flask uygulamasını çalıştırıp `5000` portunu ngrok ile internete açın. Ngrok'un HTTPS adresini Android uygulamasında ilk açılışta girin. Ngrok adresi değişirse uygulamanın sakladığı adresi Android Ayarlar > Uygulamalar > Dükkan Asistanı > Depolama > Verileri temizle yoluyla sıfırlayın; uygulamayı yeniden açıp yeni adresi girin. Verileri temizlemek cihazdaki oturumu ve kayıtlı adresi siler.

Ngrok HTTPS üzerinden erişim sağlar; Flask sunucusunu internete açmak tek başına ERP/Firebird veritabanını telefona taşımaz. Veritabanı bağlantısı ve Akınsoft otomasyon işlemleri sunucunun çalıştığı bilgisayarda kalır. Tahsilat, yazdırma ve diğer Akınsoft otomasyonları için Akınsoft'un da o bilgisayarda açık olması gerekir.
