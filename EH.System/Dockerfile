#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

#Depending on the operating system of the host machines(s) that will build or run the containers, the image specified in the FROM statement may need to be changed.
#For more information, please see https://aka.ms/containercompat

FROM mcr.microsoft.com/dotnet/aspnet:6.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["EH.System/EH.System.csproj", "EH.System/"]
COPY ["EH.Service/EH.Service.csproj", "EH.Service/"]
COPY ["EH.System.Commons/EH.System.Commons.csproj", "EH.System.Commons/"]
COPY ["EH.System.Models/EH.System.Models.csproj", "EH.System.Models/"]
COPY ["EH.Repository/EH.Repository.csproj", "EH.Repository/"]
RUN dotnet restore "./EH.System/EH.System.csproj"
COPY . .
WORKDIR "/src/EH.System"
RUN dotnet build "./EH.System.csproj" -c %BUILD_CONFIGURATION% -o /app/build

FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./EH.System.csproj" -c %BUILD_CONFIGURATION% -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "EH.System.dll"]