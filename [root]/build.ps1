az login --use-device-code
az acr login --name [acrcontainer]
docker build -f Docker\Dockerfile . -t [classicaspapp]
docker tag [classicaspapp] [acrcontainer].azurecr.io/[classicaspapp]:v1
docker push [acrcontainer].azurecr.io/[classicaspapp]:v1