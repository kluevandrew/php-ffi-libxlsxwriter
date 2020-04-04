FROM php:7.4-cli-alpine

RUN apk add --no-cache --update \
        bash \
        nano \
        libffi-dev \
        autoconf \
        g++ \
        make \
        curl \
    && rm -rf /var/cache/apk/* \
    && docker-php-ext-install ffi \
    && COMPOSER_ALLOW_SUPERUSER=1 curl -sS https://getcomposer.org/installer | php -- \
        --filename=composer \
        --install-dir=/usr/local/bin \
    && composer global require hirak/prestissimo --no-plugins --no-scripts \
    && rm -rf /root/.composer \
    && curl 'https://codeload.github.com/jmcnamara/libxlsxwriter/tar.gz/RELEASE_0.9.4' > 'libxlsxwriter.tar.gz' \
    && tar xzfv libxlsxwriter.tar.gz

RUN apk add --no-cache --update zlib-dev
RUN cd 'libxlsxwriter-RELEASE_0.9.4' \
    && make \
    && make install

RUN ls -al /usr/local/lib | grep libxlsxwriter

WORKDIR /app
CMD ["tail", "-f", "/dev/null"]
