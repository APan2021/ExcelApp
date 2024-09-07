cd excel-app-portal

docker build -t excel-app-portal .

cd ../

docker build -t excel-app .

docker-compose up -d