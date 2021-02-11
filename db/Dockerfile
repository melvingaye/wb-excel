FROM mysql:5.7

ENV MYSQL_DATABASE shipping
ENV MYSQL_ROOT_PASSWORD password
ENV MYSQL_USER mysql 
ENV MYSQL_PASSWORD password
ENV DATABASE_HOST db

EXPOSE 3306

COPY ./data.sql /docker-entrypoint-initdb.d/data.sql