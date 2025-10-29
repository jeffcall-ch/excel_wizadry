# How to Add Application Icon

The project is configured to use `app.ico` as the application icon. You need to add this file to make the icon appear in the taskbar.

## Option 1: Download a Search Icon (Recommended)

1. Visit one of these free icon sites:
   - https://icons8.com/ (search for "search" or "magnifying glass")
   - https://www.flaticon.com/ (search for "search")
   - https://www.iconfinder.com/ (search for "search")

2. Download icon in **ICO format** with multiple sizes (16x16, 32x32, 48x48, 256x256)

3. Save the file as: `c:\Users\szil\Repos\excel_wizadry\qr_search_gui_csharp\RqGui\app.ico`

4. Rebuild the application:
   ```powershell
   cd c:\Users\szil\Repos\excel_wizadry\qr_search_gui_csharp\RqGui
   dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
   ```

## Option 2: Create Icon from Image

If you have a PNG or JPG image you want to use:

1. Use an online converter:
   - https://convertio.co/png-ico/
   - https://www.icoconverter.com/

2. Upload your image and select these sizes: 16x16, 32x32, 48x48, 256x256

3. Download the ICO file

4. Save as: `c:\Users\szil\Repos\excel_wizadry\qr_search_gui_csharp\RqGui\app.ico`

5. Rebuild (see command above)

## Option 3: Use Windows Built-in Icons

You can extract icons from Windows system files:

1. Press Win+R and type: `shell:ResourceDir`
2. Look through the DLL files for icons (right-click → Properties → Icons tab)
3. Use a tool like "IconsExtract" from NirSoft to extract
4. Save as `app.ico` in the RqGui folder

## Temporary Solution: Use Default .NET Icon

If you want to proceed without a custom icon, you can remove the `<ApplicationIcon>` line from `RqGui.csproj` and it will use the default .NET application icon.

## Recommended Icon Style

For a search application, consider icons with:
- 🔍 Magnifying glass
- 📁 Folder with search
- 🔎 Search document
- 🎯 Target/search crosshair

Colors that work well: Blue, Green, or Orange on white/transparent background
