# Multi-stage Dockerfile for Gupta Brothers Shipping & Tax Invoice Suite
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /app

# Copy solution and project files first to leverage Docker layer caching
COPY GuptaBrothers.sln ./
COPY src/GuptaBrothers.Web/*.csproj ./src/GuptaBrothers.Web/
RUN dotnet restore src/GuptaBrothers.Web/GuptaBrothers.Web.csproj

# Copy remaining source code and publish
COPY . ./
RUN dotnet publish src/GuptaBrothers.Web/GuptaBrothers.Web.csproj -c Release -o /app/out

# Production runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
WORKDIR /app
COPY --from=build /app/out ./

# Expose standard container port
ENV ASPNETCORE_ENVIRONMENT=Production
ENV ASPNETCORE_URLS=http://+:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "GuptaBrothers.Web.dll"]
