.PHONY: all

all:
	dotnet restore "Wochenstatistik.csproj"
	dotnet build "Wochenstatistik.csproj" -c Release -o ./build
	dotnet publish "Wochenstatistik.csproj" -c Release -o ./publish /p:UseAppHost=false
	dotnet ./publish/Wochenstatistik.dll

bin-mac:
	dotnet restore "Wochenstatistik.csproj"
	dotnet build "Wochenstatistik.csproj" -c Release -o ./build
	dotnet publish "Wochenstatistik.csproj" -c Release -o ./my_bins/bin_mac --self-contained -r osx-x64 /p:PublishSingleFile=true

bin-win:
	dotnet restore "Wochenstatistik.csproj"
	dotnet build "Wochenstatistik.csproj" -c Release -o ./build
	dotnet publish "Wochenstatistik.csproj" -c Release -o ./my_bins/bin_win --self-contained -r win-x64  /p:PublishSingleFile=true