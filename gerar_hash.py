# gerar_hash.py
import hashlib

# IMPORTANTE: Coloque aqui a sua chave de API original, a mesma que você escolheu antes.
minha_chave_secreta = "waha-eh-show-de-bola-123"

# Este código vai gerar o hash SHA512
hash_sha512 = hashlib.sha512(minha_chave_secreta.encode('utf-8')).hexdigest()

print("--- Seu Hash SHA512 ---")
print(hash_sha512)
print("\nCopie a linha de código abaixo e use no seu comando 'docker run':")
print(f'WAHA_API_KEY=sha512:{hash_sha512}')