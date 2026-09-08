# Use PHP 8.2 with Apache
FROM php:8.2-apache

# 1. Install system dependencies: LibreOffice, Imagick, fonts, and compression libraries
RUN apt-get update && apt-get install -y \
    libmagickwand-dev \
    libpng-dev \
    libjpeg-dev \
    libfreetype6-dev \
    libzip-dev \
    libxml2-dev \
    zip \
    unzip \
    libreoffice \
    fonts-noto-cjk \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# 2. Install and enable the required PHP extensions
RUN pecl install imagick \
    && docker-php-ext-enable imagick \
    && docker-php-ext-configure gd --with-freetype --with-jpeg \
    && docker-php-ext-install gd zip xml bcmath

# 3. Enable the Apache rewrite module
RUN a2enmod rewrite

# 4. Install Composer
COPY --from=composer:latest /usr/bin/composer /usr/bin/composer

# 5. Set the working directory
WORKDIR /var/www/html

# 6. Copy the project files
COPY . .

# 7. Install PHP dependencies
RUN composer install --no-interaction --optimize-autoloader

# 8. Create a temporary directory and assign full permissions to allow storage of uploaded files and generated conversion results
RUN mkdir -p /var/www/html/temp_uploads && \
    chmod -R 777 /var/www/html/temp_uploads

# 9. Automatically adapt to Render's assigned port (Render dynamically assigns the $PORT environment variable)
RUN sed -i 's/80/${PORT}/g' /etc/apache2/sites-available/000-default.conf /etc/apache2/ports.conf

# Expose the application port
EXPOSE 80
