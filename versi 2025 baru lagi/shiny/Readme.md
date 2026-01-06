# read me

## how to build

_login_ dulu

```
docker login
```

```
docker build -t ikanx101/shiny-so-appsheet .
```

```
sudo docker tag ikanx101/shiny-so-appsheet:latest ikanx101/shiny-so-appsheet:latest
```

```
sudo docker push ikanx101/shiny-so-appsheet:latest
```

## how to run

```
docker run -p 3333:3838 -d --name so-appsheet_converter --restart unless-stopped ikanx101/shiny-so-appsheet:latest
```
