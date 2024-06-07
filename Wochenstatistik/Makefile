.PHONY: all

all:
	dotnet restore "Wochenstatistik.csproj"
	dotnet build "Wochenstatistik.csproj" -c Release -o ./build
	dotnet publish "Wochenstatistik.csproj" -c Release -o ./publish /p:UseAppHost=false
	dotnet ./publish/Wochenstatistik.dll