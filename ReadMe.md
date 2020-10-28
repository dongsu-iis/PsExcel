# PowershellからExcel操作のサンプルプログラム

## 手順

1. ClosedXMLのdllを生成する
   - [ClosedXML](https://github.com/ClosedXML/ClosedXML)のリポジトリをクローン
   - Visual Studioで`ClosedXML.sln`を開き、ClosedXMLプロジェクトをビルドする(構成はRelease)
1. dllをPowershellと同じフォルダへ配置する
   - `ClosedXML\bin\Release\net40`フォルダにある`ClosedXML.dll`と`DocumentFormat.OpenXml.dll`をPowershellと同じ場所へコピー
2. `WriteToExcel.ps1`を実行する
