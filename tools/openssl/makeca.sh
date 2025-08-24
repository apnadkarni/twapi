# Reference: https://pki-tutorial.readthedocs.io/en/latest/simple/#create-ca-certificate
echo Creating root CA directories...
mkdir -p ca/root-ca/private ca/root-ca/db crl certs
echo Initialize databases...
cp /dev/null ca/root-ca/db/root-ca.db
echo 01 > ca/root-ca/db/root-ca.crt.srl
echo 01 > ca/root-ca/db/root-ca.crl.srl
echo Creating the root CA CSR...
openssl req -new \
        -config etc/root-ca.conf \
        -out ca/root-ca.csr \
        -keyout ca/root-ca/private/root-ca.key
echo Creating the root CA self signed cert...
openssl ca -selfsign \
        -config etc/root-ca.conf \
        -in ca/root-ca.csr \
        -out ca/root-ca.crt \
        -extensions root_ca_ext

echo Creating Signing CA directories...
mkdir -p ca/signing-ca/private ca/signing-ca/db crl certs
echo Initialize Signing CA databases...
cp /dev/null ca/signing-ca/db/signing-ca.db
echo 01 > ca/signing-ca/db/signing-ca.crt.srl
echo 01 > ca/signing-ca/db/signing-ca.crl.srl
echo Creating signing CA CSR...
openssl req -new \
        -config etc/signing-ca.conf \
        -out ca/signing-ca.csr \
        -keyout ca/signing-ca/private/signing-ca.key
echo Creating signing CA certificate...
openssl ca \
        -config etc/root-ca.conf \
        -in ca/signing-ca.csr \
        -out ca/signing-ca.crt \
        -extensions signing_ca_ext

echo Creating client CSR...
openssl req -new \
        -config etc/client.conf \
        -out certs/simpleuser.csr \
        -keyout certs/simpleuser.key
echo Creating client cert...
openssl ca \
        -config etc/signing-ca.conf \
        -in certs/simpleuser.csr \
        -out certs/simpleuser.crt \
        -extensions client_ext

echo Creating server CSR...
SAN=DNS:www.simple.org \
    openssl req -new \
    -config etc/server.conf \
    -out certs/simple.org.csr \
    -keyout certs/simple.org.key

openssl ca \
        -config etc/signing-ca.conf \
        -in certs/simple.org.csr \
        -out certs/simple.org.crt \
        -extensions server_ext
