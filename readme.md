# Transforms our step report output to a file format specific to Mayo

To use in dev:
- `dotnet run path/to/excel.xlsx`

To publish:
- `dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true`

To use published version:
- drag excel file onto executable, it will generate a file in the same directory as the original file, with `-transformed-{timestamp}` appended to the name