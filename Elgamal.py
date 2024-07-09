import hashlib
def gcd(a, b):
    while a != 0:
        a, b = b % a, a
    return b

# Binh phuong va nhan
def binhPhuong(a, b, n):
    if b == 0:
        return 1
    p = binhPhuong(a, b // 2, n) % n
    if b % 2 == 0:
        return (p * p) % n
    else:
        return (p * p * a) % n

def nghichDao(a,m):
    m0 = m
    x = 1; y = 0
    if m == 1:
        return 0
    while a > 1:
        q = int(a / m)
        t = m
        m = a % m
        a = t
        t = y
        y = x - q*y
        x = t
    if x<0:
        x = x + m0
    return x
def kiemTraNguyenTo(n):
    if (n < 2):
        return 0

    for i in range(2, int(n**0.5) + 1):
        if n % i == 0:
            return 0
            break
    return 1
def sha256(message):

    # Băm thông điệp sử dụng hàm băm SHA-256
    hashed_message = hashlib.sha256(message.encode()).hexdigest()
    return hashed_message

def chuyenVBsangSo(message):
    hashed_message = sha256(message)
    # Chuyển đổi giá trị băm thành số nguyên dương
    hashed_value = int(hashed_message,16)
    return hashed_value
def thucHienKy(vb1, p, a, k, gamma):
    x = chuyenVBsangSo(vb1)
    number = nghichDao(k, p-1)
    deta = (number * (x - a * gamma)) % (p -1 )
    return deta
def kiemTraChuKy(vb2, gamma, deta, alpha, beta, p):
    beta_gamma = binhPhuong(beta, gamma, p)
    gamma_deta = binhPhuong(gamma, deta, p)
    m = (beta_gamma * gamma_deta) % p
    x = chuyenVBsangSo(vb2)
    n = binhPhuong(alpha, x, p)
    if m == n:
        return 1
    else:
        return 0