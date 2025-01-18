# Field Marker for Microsoft Word 

Microsoft Word'deki alanları (TOC, çapraz referanslar vb.) tek tıklamayla renklendirebilen bir VSTO eklentisi.

[English](#english) | [Türkçe](#türkçe)

## English

### Features

- Simple and user-friendly interface with "Marker" tab
- Instantly highlight all fields in your document
- Choose from multiple highlight colors:
  - Yellow
  - Green
  - Turquoise
  - Pink
  - Red
  - Blue
  - Dark Blue
  - Teal
  - Gray
- "No Color" option to remove highlighting
- Clear feedback messages showing number of fields processed
- Error handling for various scenarios

### Requirements

- Microsoft Word 2013 or later
- .NET Framework 4.8
- VSTO Runtime

### Installation

1. Download the Setup.zip file from the latest release
2. Extract the Setup.zip file
3. Run the ReferencedMarker.vsto file
4. If prompted about the publisher, click "Install"
5. Wait for the installation to complete
6. Restart Microsoft Word

Note: If you encounter any issues during installation, make sure you have the VSTO Runtime installed. You can download it from the Microsoft website.

### Usage

1. Open Microsoft Word
2. Go to the "Marker" tab
3. Select a color from the color picker
4. Click the color button to highlight all fields
5. Use "No Color" to remove highlighting

## Türkçe

### Özellikler

- "Marker" sekmesiyle basit ve kullanıcı dostu arayüz
- Dokümandaki tüm alanları tek tıklamayla renklendirme
- Birden fazla renk seçeneği:
  - Sarı
  - Yeşil
  - Turkuaz
  - Pembe
  - Kırmızı
  - Mavi
  - Koyu Mavi
  - Turkuaz
  - Gri
- "Renk Yok" seçeneğiyle renklendirmeyi kaldırma
- İşlenen alan sayısını gösteren bildirimler
- Çeşitli senaryolar için hata yönetimi

### Gereksinimler

- Microsoft Word 2013 veya üzeri
- .NET Framework 4.8
- VSTO Runtime

### Kurulum

1. Son sürümdeki Setup.zip dosyasını indirin
2. Setup.zip dosyasını çıkartın
3. ReferencedMarker.vsto dosyasını çalıştırın
4. Yayıncı ile ilgili bir uyarı gelirse "Yükle"ye tıklayın
5. Kurulumun tamamlanmasını bekleyin
6. Microsoft Word'ü yeniden başlatın

Not: Kurulum sırasında sorun yaşarsanız, VSTO Runtime'ın yüklü olduğundan emin olun. Microsoft web sitesinden indirebilirsiniz.

### Kullanım

1. Microsoft Word'ü açın
2. "Marker" sekmesine gidin
3. Renk seçiciden bir renk seçin
4. Tüm alanları renklendirmek için renk butonuna tıklayın
5. Renklendirmeyi kaldırmak için "Renk Yok"u kullanın

## Development

### Tech Stack

- C# (.NET Framework 4.8)
- VSTO (Visual Studio Tools for Office)
- Windows Forms

### Project Structure

- `CustomRibbon.cs`: Main functionality implementation
- `CustomRibbon.xml`: Ribbon UI definition
- `ThisAddIn.cs`: Add-in initialization
- `Resources/`: Icons and images

## License

MIT License

## Author

- **Burak Can KARA**
- GitHub: [github.com/bcankara](https://github.com/bcankara)
- Email: burakcankara@outlook.com

## Repository

[github.com/bcankara/fieldMarker](https://github.com/bcankara/fieldMarker) 