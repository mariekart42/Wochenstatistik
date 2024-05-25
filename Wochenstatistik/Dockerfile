FROM mcr.microsoft.com/dotnet/runtime:8.0 AS base
USER $APP_UID
WORKDIR /app
COPY Wochenstatistik/document/Daten_Wochenstatistik.xlsx /document/

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["Wochenstatistik/Wochenstatistik.csproj", "Wochenstatistik/"]
RUN dotnet restore "Wochenstatistik/Wochenstatistik.csproj"
COPY . .
WORKDIR "/src/Wochenstatistik"
RUN dotnet build "Wochenstatistik.csproj" -c $BUILD_CONFIGURATION -o /app/build

FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "Wochenstatistik.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Wochenstatistik.dll"]
