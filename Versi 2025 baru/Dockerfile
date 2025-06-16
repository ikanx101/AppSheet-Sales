FROM rocker/shiny

RUN apt-get update && apt-get install -y --no-install-recommends \
    pandoc \
    curl \
    gdebi-core \
    && rm -rf /var/lib/apt/lists/*

RUN curl -LO https://github.com/quarto-dev/quarto-cli/releases/download/v1.6.42/quarto-1.6.42-linux-amd64.deb
RUN gdebi --non-interactive quarto-1.6.42-linux-amd64.deb

RUN Rscript -e "install.packages(c('dplyr','reshape2','tidyr','expss'))"
RUN Rscript -e "install.packages(c('shiny','quarto','openxlsx'))"
RUN Rscript -e "install.packages(c('janitor','bslib','readxl'))"

# membuat working directory yang diperlukan
RUN mkdir -p /srv/shiny-server/dashboard && \
    chown -R shiny:shiny /srv/shiny-server

# kita copy server configuration ke dalam working directory
COPY shiny-server.conf /etc/shiny-server/shiny-server.conf

# membuat working directory yang diperlukan
RUN mkdir -p /var/log/shiny-server && \
    chown -R shiny:shiny /var/log/shiny-server

# kita copy shiny apps ke dalam working directory
COPY index.qmd /srv/shiny-server/dashboard/index.qmd

# set working directory
WORKDIR /srv/shiny-server/dashboard/

# run quarto
RUN quarto render index.qmd

USER shiny
CMD ["/usr/bin/shiny-server"]


#docker build -t ikanx101/appsheet_converter . 
#docker tag ikanx101/appsheet_converter:latest ikanx101/appsheet_converter
#docker push ikanx101/appsheet_converter:latest



