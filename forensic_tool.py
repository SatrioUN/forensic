import os
import hashlib
import time
import subprocess
import magic # Anda mungkin perlu menginstal: pip install python-magic-bin (Windows) atau pip install python-magic (Linux/macOS)
import re
import traceback
from PIL import Image, UnidentifiedImageError # pip install Pillow
from PIL.ExifTags import TAGS
from docx import Document # pip install python-docx
from geopy.geocoders import Nominatim # pip install geopy
from geopy.exc import GeocoderTimedOut, GeocoderServiceError
import sys # Ditambahkan untuk sys.exit()

# ANSI Colors
class Colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    VIOLET = '\033[95m'
    BLINK = '\033[5m'
    ORANGE = '\033[38;5;208m'
    GRAY = '\033[90m'

# Dictionary mapping wilayah (sementara, perlu data lengkap)
wilayah_dict = {
    '11': 'Aceh', '12': 'Sumatera Utara', '13': 'Sumatera Barat', '14': 'Riau',
    '15': 'Jambi', '16': 'Sumatera Selatan', '17': 'Bengkulu', '18': 'Lampung',
    '19': 'Kepulauan Bangka Belitung', '21': 'Kepulauan Riau', '31': 'DKI Jakarta',
    '32': 'Jawa Barat', '33': 'Jawa Tengah', '34': 'DI Yogyakarta', '35': 'Jawa Timur',
    '36': 'Banten', '51': 'Bali', '52': 'Nusa Tenggara Barat', '53': 'Nusa Tenggara Timur',
    '61': 'Kalimantan Barat', '62': 'Kalimantan Tengah', '63': 'Kalimantan Selatan',
    '64': 'Kalimantan Timur', '65': 'Kalimantan Utara', '71': 'Sulawesi Utara',
    '72': 'Sulawesi Tengah', '73': 'Sulawesi Selatan', '74': 'Sulawesi Tenggara',
    '75': 'Gorontalo', '76': 'Sulawesi Barat', '81': 'Maluku', '82': 'Maluku Utara',
    '91': 'Papua Barat', '94': 'Papua'
    # Tambahkan kode wilayah lainnya di sini
}

# Dictionary mapping kabupaten/kota (sementara, perlu data lengkap)
# Kunci = Kode Provinsi + Kode Kab/Kota (4 digit)
kabupaten_kota_dict = {
    '3171': 'Jakarta Selatan', '3172': 'Jakarta Timur', '3173': 'Jakarta Pusat',
    '3174': 'Jakarta Barat', '3175': 'Jakarta Utara', '3201': 'Kabupaten Bogor',
    '3271': 'Kota Bogor', '3202': 'Kabupaten Sukabumi', '3272': 'Kota Sukabumi',
    '3203': 'Kabupaten Cianjur', '3204': 'Kabupaten Bandung', '3273': 'Kota Bandung',
    '3205': 'Kabupaten Garut', '3206': 'Kabupaten Tasikmalaya', '3278': 'Kota Tasikmalaya',
    '3207': 'Kabupaten Ciamis', '3208': 'Kabupaten Kuningan', '3209': 'Kabupaten Cirebon',
    '3274': 'Kota Cirebon', '3210': 'Kabupaten Majalengka', '3211': 'Kabupaten Sumedang',
    '3212': 'Kabupaten Indramayu', '3213': 'Kabupaten Subang', '3214': 'Kabupaten Purwakarta',
    '3215': 'Kabupaten Karawang', '3216': 'Kabupaten Bekasi', '3275': 'Kota Bekasi',
    '3217': 'Kabupaten Bandung Barat', '3277': 'Kota Cimahi', '3218': 'Kabupaten Pangandaran',
    '3301': 'Kabupaten Cilacap', '3302': 'Kabupaten Banyumas', '3303': 'Kabupaten Purbalingga',
    '3304': 'Kabupaten Banjarnegara', '3305': 'Kabupaten Kebumen', '3306': 'Kabupaten Purworejo',
    '3307': 'Kabupaten Wonosobo', '3308': 'Kabupaten Magelang', '3371': 'Kota Magelang',
    '3309': 'Kabupaten Boyolali', '3310': 'Kabupaten Klaten', '3311': 'Kabupaten Sukoharjo',
    '3312': 'Kabupaten Wonogiri', '3313': 'Kabupaten Karanganyar', '3314': 'Kabupaten Sragen',
    '3315': 'Kabupaten Grobogan', '3316': 'Kabupaten Blora', '3317': 'Kabupaten Rembang',
    '3318': 'Kabupaten Pati', '3319': 'Kabupaten Kudus', '3320': 'Kabupaten Jepara',
    '3321': 'Kabupaten Demak', '3322': 'Kabupaten Semarang', '3372': 'Kota Salatiga',
    # Kota Semarang (3374) tampaknya salah tempat, seharusnya setelah Kab. Semarang
    '3374': 'Kota Semarang', # Pindah posisi, sebelumnya terdaftar sebagai Kota Surakarta
    '3323': 'Kabupaten Kendal', # Sebelumnya terdaftar sbg Kab Semarang
    '3324': 'Kabupaten Batang', # Sebelumnya terdaftar sbg Kota Salatiga
    '3325': 'Kabupaten Pekalongan', # Sebelumnya terdaftar sbg Kab Kendal
    '3375': 'Kota Pekalongan', # Sebelumnya terdaftar sbg Kab Batang
    '3326': 'Kabupaten Pemalang', # Sebelumnya terdaftar sbg Kab Pekalongan
    '3376': 'Kota Tegal', # Sebelumnya terdaftar sbg Kota Pekalongan
    '3327': 'Kabupaten Tegal', # Sebelumnya terdaftar sbg Kab Pemalang
    '3328': 'Kabupaten Brebes', # Sebelumnya terdaftar sbg Kab Tegal
    '3329': 'Kabupaten Purworejo', # Duplikat? Kode Purworejo adalah 3306
    '3373': 'Kota Surakarta', # Kode yang benar untuk Surakarta
    '3401': 'Kabupaten Kulon Progo', '3402': 'Kabupaten Bantul', '3403': 'Kabupaten Gunung Kidul',
    '3404': 'Kabupaten Sleman', '3471': 'Kota Yogyakarta', '3501': 'Kabupaten Pacitan',
    '3502': 'Kabupaten Ponorogo', '3503': 'Kabupaten Trenggalek', '3504': 'Kabupaten Tulungagung',
    '3505': 'Kabupaten Blitar', '3572': 'Kota Blitar', '3506': 'Kabupaten Kediri',
    '3571': 'Kota Kediri', '3507': 'Kabupaten Malang', '3573': 'Kota Malang',
    '3508': 'Kabupaten Lumajang', '3509': 'Kabupaten Jember', '3510': 'Kabupaten Banyuwangi',
    '3511': 'Kabupaten Bondowoso', '3512': 'Kabupaten Situbondo', '3513': 'Kabupaten Probolinggo',
    '3574': 'Kota Probolinggo', '3514': 'Kabupaten Pasuruan', '3575': 'Kota Pasuruan', # Ditambahkan
    '3515': 'Kabupaten Sidoarjo', '3578': 'Kota Surabaya', '3516': 'Kabupaten Mojokerto',
    '3576': 'Kota Mojokerto', '3517': 'Kabupaten Jombang', '3518': 'Kabupaten Nganjuk',
    '3519': 'Kabupaten Madiun', '3577': 'Kota Madiun', '3520': 'Kabupaten Magetan',
    '3521': 'Kabupaten Ngawi', '3522': 'Kabupaten Bojonegoro', '3523': 'Kabupaten Tuban',
    '3524': 'Kabupaten Lamongan', '3525': 'Kabupaten Gresik', '3526': 'Kabupaten Bangkalan',
    '3527': 'Kabupaten Sampang', '3528': 'Kabupaten Pamekasan', '3529': 'Kabupaten Sumenep',
    '3579': 'Kota Batu', # Ditambahkan
    '3601': 'Kabupaten Pandeglang', '3602': 'Kabupaten Lebak', '3603': 'Kabupaten Tangerang',
    '3671': 'Kota Tangerang', '3604': 'Kabupaten Serang', # Ditambahkan
    '3672': 'Kota Cilegon', '3673': 'Kota Serang', '3674': 'Kota Tangerang Selatan'
}

# Dictionary mapping kecamatan (sementara, perlu data lengkap)
# Kunci = Kode Provinsi + Kode Kab/Kota + Kode Kecamatan + '0' (7 digit asumsi?)
# PERHATIAN: Format kunci ini (7 digit) mungkin tidak sesuai dengan parsing NIK standar (6 digit untuk kecamatan).
# Pencarian di parse_nik() mencoba mencocokkan dengan menambahkan '0'.
kecamatan_dict = {
    '3171010': 'Kebayoran Baru', '3171020': 'Kebayoran Lama', # Urutan diperbaiki
    '3171030': 'Pesanggrahan', # Urutan diperbaiki
    '3171040': 'Cilandak', # Urutan diperbaiki
    '3171050': 'Pasar Minggu', # Urutan diperbaiki
    '3171060': 'Jagakarsa', # Urutan diperbaiki
    '3171070': 'Mampang Prapatan', # Urutan diperbaiki
    '3171080': 'Pancoran', # Urutan diperbaiki
    '3171090': 'Tebet', # Urutan diperbaiki
    '3171100': 'Setiabudi', # Urutan diperbaiki
    '3172010': 'Matraman', '3172020': 'Pulo Gadung', '3172030': 'Jatinegara',
    '3172040': 'Kramat Jati', # Urutan diperbaiki
    '3172050': 'Pasar Rebo', # Urutan diperbaiki
    '3172060': 'Cakung', # Urutan diperbaiki
    '3172070': 'Duren Sawit', # Urutan diperbaiki
    '3172080': 'Makasar', # Urutan diperbaiki
    '3172090': 'Ciracas', # Urutan diperbaiki
    '3172100': 'Cipayung', # Urutan diperbaiki
    '3173010': 'Gambir', '3173020': 'Tanah Abang', # Urutan diperbaiki
    '3173030': 'Menteng', # Urutan diperbaiki
    '3173040': 'Senen', # Urutan diperbaiki
    '3173050': 'Cempaka Putih', # Urutan diperbaiki
    '3173060': 'Johar Baru', # Urutan diperbaiki
    '3173070': 'Kemayoran', # Urutan diperbaiki
    '3173080': 'Sawah Besar', # Urutan diperbaiki
    '3174010': 'Cengkareng', # Urutan diperbaiki
    '3174020': 'Grogol Petamburan', '3174030': 'Taman Sari', # Urutan diperbaiki
    '3174040': 'Tambora', # Ditambahkan (kode umum, perlu verifikasi)
    '3174050': 'Kebon Jeruk', # Urutan diperbaiki
    '3174060': 'Kalideres', # Urutan diperbaiki
    '3174070': 'Palmerah', # Urutan diperbaiki
    '3174080': 'Kembangan', # Urutan diperbaiki
    # '3174080': 'Curug', # Curug ada di Kab. Tangerang (3603) atau Kota Tangerang (3671)
    '3175010': 'Penjaringan', '3175020': 'Tanjung Priok', '3175030': 'Koja',
    '3175040': 'Cilincing', '3175050': 'Pademangan', '3175060': 'Kelapa Gading',
    # '3175070': 'Sunter', # Sunter adalah kelurahan di Tanjung Priok
    # Kode Kecamatan Bogor (Kabupaten)
    '3201010': 'Cibinong', '3201020': 'Gunung Putri', '3201030': 'Citeureup',
    '3201040': 'Sukaraja', # Urutan diperbaiki
    '3201050': 'Babakan Madang', # Urutan diperbaiki
    '3201060': 'Jonggol', # Urutan diperbaiki
    '3201070': 'Cileungsi', # Urutan diperbaiki
    '3201080': 'Cariu', # Ditambahkan (kode umum)
    '3201090': 'Sukamakmur', # Urutan diperbaiki
    '3201100': 'Parung', # Urutan diperbaiki
    '3201110': 'Gunung Sindur', # Urutan diperbaiki
    '3201120': 'Kemang', # Urutan diperbaiki
    '3201130': 'Bojonggede', # Urutan diperbaiki
    '3201140': 'Leuwiliang', # Ditambahkan (kode umum)
    '3201150': 'Ciampea', # Urutan diperbaiki
    '3201160': 'Cibungbulang', # Ditambahkan (kode umum)
    '3201170': 'Pamijahan', # Ditambahkan (kode umum)
    '3201180': 'Rumpin', # Urutan diperbaiki
    '3201190': 'Jasinga', # Ditambahkan (kode umum)
    '3201200': 'Parung Panjang', # Urutan diperbaiki
    '3201210': 'Nanggung', # Ditambahkan (kode umum)
    '3201220': 'Cigudeg', # Ditambahkan (kode umum)
    '3201230': 'Tenjo', # Ditambahkan (kode umum)
    '3201240': 'Ciawi', # Urutan diperbaiki
    '3201250': 'Cisarua', # Urutan diperbaiki
    '3201260': 'Megamendung', # Urutan diperbaiki
    '3201270': 'Caringin', # Urutan diperbaiki ('caringin' -> 'Caringin')
    '3201280': 'Cijeruk', # Urutan diperbaiki
    '3201290': 'Ciomas', # Urutan diperbaiki
    '3201300': 'Dramaga', # Urutan diperbaiki
    '3201310': 'Tamansari', # Urutan diperbaiki
    '3201320': 'Klapanunggal', # Ditambahkan (kode umum)
    '3201330': 'Ciseeng', # Urutan diperbaiki
    '3201340': 'Ranca Bungur', # Urutan diperbaiki (' Ranca Bungur' -> 'Ranca Bungur')
    '3201350': 'Sukajaya', # Ditambahkan (kode umum)
    '3201360': 'Tajurhalang', # Urutan diperbaiki
    '3201370': 'Cigombong', # Urutan diperbaiki
    '3201380': 'Leuwisadeng', # Ditambahkan (kode umum)
    '3201390': 'Tenjolaya', # Ditambahkan (kode umum)
    '3201400': 'Tanjungsari', # Ditambahkan (kode umum)
    # Kode Kecamatan Bogor (Kota)
    '3271010': 'Bogor Selatan', '3271020': 'Bogor Timur', '3271030': 'Bogor Utara', # Urutan diperbaiki
    '3271040': 'Bogor Tengah', # Urutan diperbaiki
    '3271050': 'Bogor Barat', # Urutan diperbaiki
    '3271060': 'Tanah Sareal', # Urutan diperbaiki
    # Kode Kecamatan Cilacap
    '3301010': 'Kedungreja', # Urutan diperbaiki
    '3301020': 'Kesugihan', # Urutan diperbaiki
    '3301030': 'Adipala', # Urutan diperbaiki
    '3301040': 'Binangun', # Urutan diperbaiki
    '3301050': 'Nusawungu', # Urutan diperbaiki
    '3301060': 'Kroya', # Urutan diperbaiki
    '3301070': 'Maos', # Urutan diperbaiki
    '3301080': 'Jeruklegi', # Urutan diperbaiki
    '3301090': 'Kawunganten', # Urutan diperbaiki (' Kawunganten' -> 'Kawunganten')
    '3301100': 'Gandrungmangu', # Urutan diperbaiki
    '3301110': 'Sidareja', # Urutan diperbaiki (' Sidareja' -> 'Sidareja')
    '3301120': 'Karangpucung', # Urutan diperbaiki (' Karangpucung' -> 'Karangpucung')
    '3301130': 'Cimanggu', # Urutan diperbaiki
    '3301140': 'Majenang', # Urutan diperbaiki
    '3301150': 'Wanareja', # Urutan diperbaiki
    '3301160': 'Dayeuhluhur', # Urutan diperbaiki
    '3301170': 'Sampang', # Urutan diperbaiki
    '3301180': 'Cipari', # Urutan diperbaiki
    '3301190': 'Patimuan', # Urutan diperbaiki
    '3301200': 'Bantarsari', # Urutan diperbaiki
    '3301210': 'Cilacap Selatan', # Urutan diperbaiki
    '3301220': 'Cilacap Tengah', # Urutan diperbaiki
    '3301230': 'Cilacap Utara', # Urutan diperbaiki
    '3301240': 'Kampung Laut', # Ditambahkan
    # Data lainnya perlu ditambahkan atau diperbaiki
}


def clear_screen():
    """Membersihkan layar terminal."""
    os.system('cls' if os.name == 'nt' else 'clear')
    # Hapus pesan blink agar tidak mengganggu
    # print(f"{Colors.VIOLET}{Colors.BLINK}Membersihkan layar...{Colors.ENDC}")
    # time.sleep(0.5) # Hapus delay agar lebih cepat

def display_header(title):
    """Menampilkan header dengan judul tertentu."""
    print("=" * 60)
    print(f"{Colors.HEADER}{Colors.BOLD}      ██████╗ ███████╗██╗  ██╗██████╗  ██████╗{Colors.ENDC}")
    print(f"{Colors.HEADER}{Colors.BOLD}      ██╔══██╗██╔════╝██║  ██║██╔══██╗██╔════╝{Colors.ENDC}")
    print(f"{Colors.HEADER}{Colors.BOLD}      ██████╔╝█████╗  ██║  ██║██████╔╝██║ ███╗{Colors.ENDC}")
    print(f"{Colors.HEADER}{Colors.BOLD}      ██╔══██╗██╔══╝   ╚██╗██╔╝██╔══██╗██║  ██║{Colors.ENDC}")
    print(f"{Colors.HEADER}{Colors.BOLD}      ██║  ██║███████╗  ╚████╔╝ ██║  ██║╚██████╔╝{Colors.ENDC}")
    print(f"{Colors.HEADER}{Colors.BOLD}      ╚═╝  ╚═╝╚══════╝   ╚═══╝  ╚═╝  ╚═╝ ╚═════╝ {Colors.ENDC}")
    print("=" * 60)
    print(f"{Colors.OKBLUE}          UBSIBER Forensic Tool - Owner: Satrio.U.N{Colors.ENDC}")
    print("=" * 60)
    print(f"{Colors.OKCYAN}{Colors.BOLD}                  {title} {Colors.ENDC}")
    print("=" * 60)

def calculate_md5(file_path):
    """Menghitung nilai hash MD5 dari file yang diberikan."""
    # display_header("[MENGHITUNG MD5 HASH]") # Header dipanggil di fungsi utama
    print(f"  Path: {Colors.BOLD}{file_path}{Colors.ENDC}")
    try:
        hasher = hashlib.md5()
        with open(file_path, 'rb') as file:
            total_size = os.path.getsize(file_path)
            bytes_read = 0
            start_time = time.time()
            while chunk := file.read(8192): # Tingkatkan ukuran chunk
                hasher.update(chunk)
                bytes_read += len(chunk)
                elapsed_time = time.time() - start_time
                if total_size > 0 : # Hindari ZeroDivisionError
                    percent_complete = (bytes_read / total_size) * 100
                    if elapsed_time > 0:
                        speed = (bytes_read / 1024) / elapsed_time # Kecepatan dalam KB/s
                        print(f"\r{Colors.ORANGE}  Menghitung MD5: {percent_complete:.2f}% | {speed:.2f} KB/s", end="")
                    else:
                        print(f"\r{Colors.ORANGE}  Menghitung MD5: {percent_complete:.2f}%", end="")
                else:
                     print(f"\r{Colors.ORANGE}  Menghitung MD5...", end="")


        # Hapus baris progress setelah selesai
        print("\r" + " " * 70 + "\r", end="") # Bersihkan baris
        md5_hash = hasher.hexdigest()
        print(f"{Colors.OKGREEN}  MD5 Hash selesai dihitung.{Colors.ENDC}")
        return md5_hash
    except FileNotFoundError:
        print(f"{Colors.FAIL}  [ERROR] File tidak ditemukan: {Colors.BOLD}{file_path}{Colors.ENDC}")
        return "[ERROR] File tidak ditemukan"
    except PermissionError:
        print(f"{Colors.FAIL}  [ERROR] Tidak ada izin untuk membaca file: {Colors.BOLD}{file_path}{Colors.ENDC}")
        return "[ERROR] Izin ditolak"
    except Exception as e:
        print(f"{Colors.FAIL}  [ERROR] Kesalahan saat menghitung MD5: {e}{Colors.ENDC}")
        return f"[ERROR] {e}"

def get_location_from_exif(exif_data):
    """Mengekstrak koordinat GPS dari data EXIF."""
    # display_header("[MENDAPATKAN LOKASI DARI EXIF]") # Header dipanggil di fungsi utama
    gps_info = None
    # Cari tag GPSInfo
    for k, v in exif_data.items():
        if TAGS.get(k) == "GPSInfo":
            gps_info = v
            break

    if gps_info:
        # print(f"Raw GPS Info: {gps_info}") # Debugging
        def _convert_to_degrees(value):
            """Konversi DMS tuple ke derajat desimal."""
            # Pastikan value adalah tuple/list dengan 3 elemen numerik
            if isinstance(value, (list, tuple)) and len(value) == 3:
                try:
                    d = float(value[0])
                    m = float(value[1])
                    s = float(value[2])
                    return d + (m / 60.0) + (s / 3600.0)
                except (ValueError, TypeError):
                     return None # Gagal konversi
            # Kadang nilai bisa berupa satu angka jika sudah desimal
            elif isinstance(value, (int, float)):
                return float(value)
            return None

        lat = lon = None
        gps_latitude = gps_info.get(2)
        gps_latitude_ref = gps_info.get(1) # 'N' atau 'S'
        gps_longitude = gps_info.get(4)
        gps_longitude_ref = gps_info.get(3) # 'E' atau 'W'

        # print(f"Lat Raw: {gps_latitude}, Ref: {gps_latitude_ref}") # Debugging
        # print(f"Lon Raw: {gps_longitude}, Ref: {gps_longitude_ref}") # Debugging

        if gps_latitude and gps_latitude_ref and gps_longitude and gps_longitude_ref:
            # Referensi kadang berupa bytes, kadang string
            if isinstance(gps_latitude_ref, bytes):
                gps_latitude_ref = gps_latitude_ref.decode('utf-8')
            if isinstance(gps_longitude_ref, bytes):
                gps_longitude_ref = gps_longitude_ref.decode('utf-8')

            lat_deg = _convert_to_degrees(gps_latitude)
            lon_deg = _convert_to_degrees(gps_longitude)

            if lat_deg is not None and lon_deg is not None:
                 lat = lat_deg if gps_latitude_ref == 'N' else -lat_deg
                 lon = lon_deg if gps_longitude_ref == 'E' else -lon_deg

        if lat is not None and lon is not None:
            print(f"{Colors.OKGREEN}  Koordinat GPS dari EXIF ditemukan.{Colors.ENDC}")
            return lat, lon
        else:
            print(f"{Colors.WARNING}  Info GPS di EXIF tidak lengkap atau format tidak dikenal.{Colors.ENDC}")
            return None, None
    else:
        print(f"{Colors.WARNING}  Tidak ada tag GPSInfo di EXIF.{Colors.ENDC}")
        return None, None

def geocode_location(latitude, longitude):
    """Melakukan reverse geocoding untuk mendapatkan informasi lokasi dari koordinat."""
    # display_header("[GEOCODING LOKASI]") # Header dipanggil di fungsi utama
    print(f"  Mencoba Geocoding untuk Latitude: {latitude:.6f}, Longitude: {longitude:.6f}")
    geolocator = Nominatim(user_agent="ubsi_forensic_tool_v1", timeout=20) # User agent unik & timeout lebih lama
    try:
        # Coba dengan bahasa Indonesia dulu
        location = geolocator.reverse(f"{latitude}, {longitude}", exactly_one=True, language='id', timeout=15)
        if not location:
             # Jika gagal, coba dengan bahasa default (Inggris)
             print(f"{Colors.WARNING}  Gagal geocoding dengan bahasa 'id', mencoba bahasa default...{Colors.ENDC}")
             time.sleep(1) # Beri jeda sebelum request lagi
             location = geolocator.reverse(f"{latitude}, {longitude}", exactly_one=True, timeout=15)

        if location:
            print(f"{Colors.OKGREEN}  Geocoding berhasil.{Colors.ENDC}")
            return location.address
        else:
             print(f"{Colors.WARNING}  Tidak ada hasil geocoding ditemukan.{Colors.ENDC}")
             return None
    except GeocoderTimedOut:
        print(f"{Colors.WARNING}  [PERINGATAN] Geocoding time out. Server mungkin sibuk atau koneksi lambat.{Colors.ENDC}")
        return None
    except GeocoderServiceError as e:
        print(f"{Colors.WARNING}  [PERINGATAN] Geocoding service error: {e}. Coba lagi nanti.{Colors.ENDC}")
        return None
    except Exception as e:
        print(f"{Colors.FAIL}  [ERROR] Kesalahan Geocoding tidak terduga: {e}{Colors.ENDC}")
        return None

def run_exiftool(image_path):
    """Menjalankan ExifTool pada path gambar yang diberikan dan mengembalikan outputnya."""
    # display_header("[MENJALANKAN EXIFTOOL]") # Header dipanggil di fungsi utama
    print(f"  Menjalankan ExifTool untuk: {Colors.BOLD}{image_path}{Colors.ENDC}")
    exiftool_path = "exiftool" # Asumsi ada di PATH
    try:
        # Cek apakah file ada sebelum menjalankan
        if not os.path.exists(image_path):
             return f"{Colors.FAIL}[ERROR] File tidak ditemukan: {Colors.BOLD}{image_path}{Colors.ENDC}"

        # Gunakan -G untuk menampilkan grup tag, -s untuk nama tag ringkas
        result = subprocess.run([exiftool_path, "-G", "-s", image_path],
                                capture_output=True, text=True, check=False, # check=False agar kita bisa tangani error
                                encoding='utf-8', errors='ignore') # Coba tangani encoding

        if result.returncode != 0:
             # Cek apakah error karena format file tidak didukung atau masalah lain
             if "File format error" in result.stderr or "Not a valid" in result.stderr:
                  return f"{Colors.WARNING}[PERINGATAN] ExifTool tidak dapat memproses format file ini atau file rusak.\nStderr: {result.stderr.strip()}{Colors.ENDC}"
             else:
                  return f"{Colors.FAIL}[ERROR] Eksekusi ExifTool gagal (kode: {result.returncode}):\n{Colors.BOLD}{result.stderr.strip()}{Colors.ENDC}"

        print(f"{Colors.OKGREEN}  ExifTool berhasil dijalankan.{Colors.ENDC}")
        return result.stdout
    except FileNotFoundError:
        return f"{Colors.WARNING}[PERINGATAN] ExifTool tidak ditemukan. Pastikan sudah terinstal dan dapat diakses dari PATH sistem Anda.{Colors.ENDC}"
    except subprocess.TimeoutExpired:
         return f"{Colors.FAIL}[ERROR] ExifTool timeout saat memproses file.{Colors.ENDC}"
    except Exception as e:
        return f"{Colors.FAIL}[ERROR] Kesalahan tak terduga saat menjalankan ExifTool: {e}{Colors.ENDC}"

def analyze_image(image_path):
    """Menganalisis gambar yang diberikan dan menampilkan detailnya (basic)."""
    display_header("[ANALISIS GAMBAR (Basic)]")
    print(f"  Menganalisis: {Colors.BOLD}{image_path}{Colors.ENDC}\n")

    if not os.path.exists(image_path):
        print(f"{Colors.FAIL}[ERROR] Gambar tidak ditemukan: {Colors.BOLD}{image_path}{Colors.ENDC}")
        return

    try:
        mime = magic.from_file(image_path, mime=True)
        print(f"  {Colors.OKBLUE}Informasi Dasar:{Colors.ENDC}")
        print(f"  MIME Type\t\t: {Colors.BOLD}{mime}{Colors.ENDC}")

        # Informasi file sistem
        try:
            file_stat = os.stat(image_path)
            file_size = file_stat.st_size
            creation_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(file_stat.st_ctime))
            modified_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(file_stat.st_mtime))
            accessed_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(file_stat.st_atime))
            print(f"  Ukuran File\t\t: {Colors.BOLD}{file_size} bytes{Colors.ENDC}")
            print(f"  Waktu Dibuat (OS)\t: {Colors.BOLD}{creation_time}{Colors.ENDC}")
            print(f"  Waktu Dimodifikasi (OS): {Colors.BOLD}{modified_time}{Colors.ENDC}")
            print(f"  Waktu Diakses (OS)\t: {Colors.BOLD}{accessed_time}{Colors.ENDC}")
        except Exception as e_stat:
            print(f"{Colors.WARNING}  [INFO] Tidak dapat membaca atribut waktu file sistem: {e_stat}{Colors.ENDC}")

        # Hitung Hash
        md5_hash = calculate_md5(image_path)
        if not md5_hash.startswith("[ERROR]"):
            print(f"  MD5 Hash\t\t: {Colors.BOLD}{md5_hash}{Colors.ENDC}")
        else:
             print(f"  MD5 Hash\t\t: {Colors.FAIL}{md5_hash}{Colors.ENDC}")

        # Analisis dengan Pillow
        try:
            with Image.open(image_path) as img:
                width, height = img.size
                img_format = img.format
                color_mode = img.mode
                print(f"  Format Gambar\t: {Colors.BOLD}{img_format}{Colors.ENDC}")
                print(f"  Dimensi\t\t: {Colors.BOLD}{width} x {height}{Colors.ENDC}")
                print(f"  Mode Warna\t\t: {Colors.BOLD}{color_mode}{Colors.ENDC}")

                # Dapatkan data EXIF menggunakan Pillow
                exif_data_pil = {}
                try:
                     exif_raw = img._getexif()
                     if exif_raw:
                          for tag, value in exif_raw.items():
                               tag_name = TAGS.get(tag, tag)
                               exif_data_pil[tag_name] = value
                          print(f"\n  {Colors.OKGREEN}Data EXIF (dari Pillow):{Colors.ENDC}")
                          relevant_tags = ['DateTimeOriginal', 'Make', 'Model', 'Software', 'Orientation', 'XResolution', 'YResolution', 'ResolutionUnit']
                          found_exif = False
                          for tag in relevant_tags:
                               if tag in exif_data_pil:
                                    print(f"  {tag}\t: {Colors.BOLD}{exif_data_pil[tag]}{Colors.ENDC}")
                                    found_exif = True
                          if not found_exif:
                               print(f"  {Colors.GRAY}Tidak ada tag EXIF relevan yang umum ditemukan.{Colors.ENDC}")

                          # Cek Lokasi dari EXIF Pillow
                          latitude, longitude = get_location_from_exif(exif_data_pil)
                          if latitude is not None and longitude is not None:
                              print(f"\n  {Colors.OKCYAN}Informasi Lokasi (GPS dari Pillow):{Colors.ENDC}")
                              print(f"  Latitude\t\t: {Colors.BOLD}{latitude:.6f}{Colors.ENDC}")
                              print(f"  Longitude\t\t: {Colors.BOLD}{longitude:.6f}{Colors.ENDC}")
                              # Format URL Google Maps yang benar
                              maps_url = f"https://www.google.com/maps?q={latitude},{longitude}"
                              print(f"  Google Maps\t\t: {Colors.UNDERLINE}{Colors.BOLD}{maps_url}{Colors.ENDC}")

                              location = geocode_location(latitude, longitude)
                              if location:
                                  print(f"  Perkiraan Lokasi\t: {Colors.BOLD}{location}{Colors.ENDC}")
                              else:
                                  print(f"  Perkiraan Lokasi\t: {Colors.WARNING}Tidak dapat mengambil nama lokasi dari koordinat.{Colors.ENDC}")
                          # else:
                              # Pesan sudah dicetak di dalam get_location_from_exif()

                     else:
                          print(f"\n  {Colors.WARNING}Tidak ada data EXIF ditemukan oleh Pillow.{Colors.ENDC}")

                except AttributeError:
                    print(f"\n  {Colors.WARNING}Tidak ada metode _getexif (mungkin bukan format yang didukung EXIF oleh Pillow).{Colors.ENDC}")
                except Exception as e_exif:
                    print(f"{Colors.FAIL}  [ERROR] Kesalahan saat membaca data EXIF Pillow: {e_exif}{Colors.ENDC}")

        except UnidentifiedImageError:
            print(f"{Colors.FAIL}[ERROR] Pillow tidak dapat mengidentifikasi format file gambar: {Colors.BOLD}{image_path}{Colors.ENDC}")
        except Exception as e_pil:
             print(f"{Colors.FAIL}[ERROR] Kesalahan saat memproses gambar dengan Pillow: {e_pil}{Colors.ENDC}")

    except FileNotFoundError: # Double check, should be caught earlier
        print(f"{Colors.FAIL}[ERROR] Gambar tidak ditemukan: {Colors.BOLD}{image_path}{Colors.ENDC}")
    except Exception as e:
        print(f"{Colors.FAIL}[ERROR] Kesalahan umum saat menganalisis gambar: {e}{Colors.ENDC}")

def analyze_image_exiftool(image_path):
    """Menganalisis gambar menggunakan ExifTool."""
    display_header("[ANALISIS GAMBAR (ExifTool)]")
    print(f"  Menganalisis: {Colors.BOLD}{image_path}{Colors.ENDC}\n")
    exiftool_output = run_exiftool(image_path)
    # Cek apakah outputnya berisi pesan error dari run_exiftool
    if "[ERROR]" in exiftool_output or "[PERINGATAN]" in exiftool_output:
         print(exiftool_output) # Tampilkan pesan error/warning
    elif exiftool_output:
        print(f"{Colors.OKGREEN}--- Output ExifTool ---{Colors.ENDC}")
        # Bersihkan sedikit output dan beri warna
        lines = exiftool_output.strip().split('\n')
        max_len = 0
        try:
             # Cari panjang kunci terpanjang untuk alignment
             key_value_pairs = [line.split(':', 1) for line in lines if ':' in line]
             if key_value_pairs:
                  max_len = max(len(pair[0].strip()) for pair in key_value_pairs)

             for line in lines:
                  parts = line.split(':', 1)
                  if len(parts) == 2:
                       key = parts[0].strip()
                       value = parts[1].strip()
                       # Tambahkan padding agar ':' lurus
                       padding = " " * (max_len - len(key))
                       print(f"  {Colors.OKBLUE}{key}{padding} : {Colors.BOLD}{value}{Colors.ENDC}")
                  else:
                       print(f"  {Colors.GRAY}{line}{Colors.ENDC}") # Baris tanpa ':' (mungkin header grup)

        except Exception as e_fmt:
             print(f"{Colors.WARNING}Gagal format output ExifTool, menampilkan mentah:{e_fmt}{Colors.ENDC}")
             print(exiftool_output) # Fallback jika parsing gagal

        print(f"{Colors.OKGREEN}--- Akhir Output ExifTool ---{Colors.ENDC}")
    else:
         print(f"{Colors.WARNING}Tidak ada output yang dihasilkan oleh ExifTool.{Colors.ENDC}")


def hash_password(password):
    """Mengembalikan nilai hash SHA-256 dari password yang diberikan."""
    display_header("[HASH PASSWORD (SHA-256)]")
    # Jangan tampilkan password asli di log/output jika sensitif
    # print(f"  Password Asli\t: {Colors.BOLD}{password}{Colors.ENDC}")
    print(f"  Menghitung hash SHA-256 untuk password yang diberikan...")
    time.sleep(0.3) # Sedikit delay kosmetik
    try:
        hashed = hashlib.sha256(password.encode('utf-8')).hexdigest()
        print(f"{Colors.OKGREEN}  Perhitungan SHA-256 selesai.{Colors.ENDC}")
        print(f"  SHA-256 Hash\t: {Colors.BOLD}{hashed}{Colors.ENDC}")
    except Exception as e:
         print(f"{Colors.FAIL}  [ERROR] Gagal menghitung hash: {e}{Colors.ENDC}")


def analyze_document(doc_path):
    """Mengekstrak dan menampilkan teks dari dokumen Word."""
    display_header("[ANALISIS DOKUMEN (.docx)]")
    print(f"  Menganalisis: {Colors.BOLD}{doc_path}{Colors.ENDC}\n")

    if not os.path.exists(doc_path):
        print(f"{Colors.FAIL}[ERROR] Dokumen tidak ditemukan: {Colors.BOLD}{doc_path}{Colors.ENDC}")
        return
    if not doc_path.lower().endswith('.docx'):
         print(f"{Colors.WARNING}[PERINGATAN] File mungkin bukan format .docx. Hasil mungkin tidak akurat.{Colors.ENDC}")

    try:
        # Info Dasar File
        mime = magic.from_file(doc_path, mime=True)
        print(f"  {Colors.OKBLUE}Informasi Dasar:{Colors.ENDC}")
        print(f"  MIME Type\t\t: {Colors.BOLD}{mime}{Colors.ENDC}")

        try:
            file_stat = os.stat(doc_path)
            file_size = file_stat.st_size
            creation_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(file_stat.st_ctime))
            modified_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(file_stat.st_mtime))
            accessed_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(file_stat.st_atime))
            print(f"  Ukuran File\t\t: {Colors.BOLD}{file_size} bytes{Colors.ENDC}")
            print(f"  Waktu Dibuat (OS)\t: {Colors.BOLD}{creation_time}{Colors.ENDC}")
            print(f"  Waktu Dimodifikasi (OS): {Colors.BOLD}{modified_time}{Colors.ENDC}")
            print(f"  Waktu Diakses (OS)\t: {Colors.BOLD}{accessed_time}{Colors.ENDC}")
        except Exception as e_stat:
            print(f"{Colors.WARNING}  [INFO] Tidak dapat membaca atribut waktu file sistem: {e_stat}{Colors.ENDC}")

        # Hash
        md5_hash = calculate_md5(doc_path)
        if not md5_hash.startswith("[ERROR]"):
            print(f"  MD5 Hash\t\t: {Colors.BOLD}{md5_hash}{Colors.ENDC}")
        else:
            print(f"  MD5 Hash\t\t: {Colors.FAIL}{md5_hash}{Colors.ENDC}")

        # Analisis Konten dengan python-docx
        print(f"\n  {Colors.OKGREEN}Metadata & Konten (dari python-docx):{Colors.ENDC}")
        doc = Document(doc_path)

        # Metadata Inti
        cp = doc.core_properties
        print(f"  Judul\t\t\t: {Colors.BOLD}{cp.title or 'Tidak Tersedia'}{Colors.ENDC}")
        print(f"  Pembuat (Author)\t: {Colors.BOLD}{cp.author or 'Tidak Tersedia'}{Colors.ENDC}")
        print(f"  Subjek\t\t: {Colors.BOLD}{cp.subject or 'Tidak Tersedia'}{Colors.ENDC}")
        print(f"  Kata Kunci\t\t: {Colors.BOLD}{cp.keywords or 'Tidak Tersedia'}{Colors.ENDC}")
        print(f"  Terakhir Dimodifikasi Oleh: {Colors.BOLD}{cp.last_modified_by or 'Tidak Tersedia'}{Colors.ENDC}")
        print(f"  Revisi\t\t: {Colors.BOLD}{cp.revision or 'Tidak Tersedia'}{Colors.ENDC}")
        try:
             # Tanggal dari metadata mungkin lebih akurat dari OS
             created_date = cp.created.strftime('%Y-%m-%d %H:%M:%S') if cp.created else "Tidak Tersedia"
             modified_date = cp.modified.strftime('%Y-%m-%d %H:%M:%S') if cp.modified else "Tidak Tersedia"
             print(f"  Dibuat (Metadata)\t: {Colors.BOLD}{created_date}{Colors.ENDC}")
             print(f"  Dimodifikasi (Metadata): {Colors.BOLD}{modified_date}{Colors.ENDC}")
        except AttributeError:
             print(f"  {Colors.WARNING}Tanggal metadata tidak dapat diakses (mungkin versi docx lama).{Colors.ENDC}")
        except ValueError:
             print(f"  {Colors.WARNING}Format tanggal metadata tidak valid.{Colors.ENDC}")


        # Jumlah Paragraf (bukan halaman)
        num_paragraphs = len(doc.paragraphs)
        print(f"  Jumlah Paragraf\t: {Colors.BOLD}{num_paragraphs}{Colors.ENDC}")
        # Jumlah Tabel
        num_tables = len(doc.tables)
        print(f"  Jumlah Tabel\t\t: {Colors.BOLD}{num_tables}{Colors.ENDC}")


        print(f"\n  {Colors.OKGREEN}Pratinjau Konten Teks (5 paragraf pertama):{Colors.ENDC}")
        if not doc.paragraphs:
             print(f"  {Colors.GRAY}(Tidak ada paragraf teks yang terdeteksi){Colors.ENDC}")
        else:
             for i, para in enumerate(doc.paragraphs):
                 if i < 5:
                     # Ambil beberapa kata pertama saja untuk pratinjau singkat
                     preview_text = " ".join(para.text.split()[:20])
                     if len(para.text.split()) > 20:
                          preview_text += "..."
                     print(f"  {Colors.GRAY}[{i+1}] {preview_text}{Colors.ENDC}")
                     time.sleep(0.1) # Sedikit delay
                 else:
                     print(f"  {Colors.WARNING}... (lebih banyak paragraf tidak ditampilkan){Colors.ENDC}")
                     break

        # Informasi Lokasi biasanya tidak ada di docx
        print(f"\n  {Colors.GRAY}[INFO] Informasi lokasi GPS umumnya tidak tersedia dalam metadata file .docx.{Colors.ENDC}")

    except ImportError:
        print(f"{Colors.FAIL}[ERROR] Library {Colors.BOLD}python-docx{Colors.ENDC}{Colors.FAIL} tidak terinstal. Silakan instal dengan: {Colors.BOLD}pip install python-docx{Colors.ENDC}")
    except FileNotFoundError: # Double check
        print(f"{Colors.FAIL}[ERROR] Dokumen tidak ditemukan: {Colors.BOLD}{doc_path}{Colors.ENDC}")
    except Exception as e: # Tangkap error spesifik jika mungkin (misal: PackageNotFoundError jika file bukan zip)
        print(f"{Colors.FAIL}[ERROR] Kesalahan saat menganalisis dokumen: {e}{Colors.ENDC}")
        print(f"{Colors.FAIL} Pastikan file adalah format .docx yang valid dan tidak rusak.{Colors.ENDC}")


def analyze_string(input_string):
    """Menganalisis string yang diberikan (basic)."""
    display_header("[ANALISIS STRING (Basic)]")
    print(f"  Input\t\t: {Colors.BOLD}{input_string[:100]}{'...' if len(input_string)>100 else ''}{Colors.ENDC}") # Tampilkan sebagian jika panjang
    try:
        encoded_string = input_string.encode('utf-8')
        md5_hash = hashlib.md5(encoded_string).hexdigest()
        sha256_hash = hashlib.sha256(encoded_string).hexdigest()
        print(f"  Panjang\t\t: {Colors.BOLD}{len(input_string)} karakter{Colors.ENDC}")
        # print(f"  Tipe\t\t: {Colors.BOLD}{type(input_string)}{Colors.ENDC}") # Kurang informatif
        print(f"  MD5 Hash\t\t: {Colors.BOLD}{md5_hash}{Colors.ENDC}")
        print(f"  SHA-256 Hash\t: {Colors.BOLD}{sha256_hash}{Colors.ENDC}")
        # Cek encoding sederhana
        try:
             input_string.encode('ascii')
             print(f"  Encoding\t\t: {Colors.BOLD}Kompatibel ASCII{Colors.ENDC}")
        except UnicodeEncodeError:
             print(f"  Encoding\t\t: {Colors.BOLD}Non-ASCII (kemungkinan UTF-8 atau lainnya){Colors.ENDC}")

    except Exception as e:
         print(f"{Colors.FAIL}  [ERROR] Gagal menganalisis string: {e}{Colors.ENDC}")


def analyze_string_advanced(input_string):
    """Menganalisis string untuk pola-pola menarik (IP, URL, Email)."""
    display_header("[ANALISIS STRING (Lanjutan)]")
    print(f"  Input\t\t: {Colors.BOLD}{input_string[:100]}{'...' if len(input_string)>100 else ''}{Colors.ENDC}") # Tampilkan sebagian
    try:
        encoded_string = input_string.encode('utf-8')
        md5_hash = hashlib.md5(encoded_string).hexdigest()
        sha256_hash = hashlib.sha256(encoded_string).hexdigest()
        print(f"  Panjang\t\t: {Colors.BOLD}{len(input_string)} karakter{Colors.ENDC}")
        print(f"  MD5 Hash\t\t: {Colors.BOLD}{md5_hash}{Colors.ENDC}")
        print(f"  SHA-256 Hash\t: {Colors.BOLD}{sha256_hash}{Colors.ENDC}")

        # Regex yang sedikit lebih baik
        # IP v4
        ip_pattern = r'\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
        # URL (lebih permisif)
        url_pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
        # Email (umum)
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

        ip_addresses = re.findall(ip_pattern, input_string)
        urls = re.findall(url_pattern, input_string)
        email_addresses = re.findall(email_pattern, input_string)

        found = False
        if ip_addresses:
            found = True
            print(f"\n  {Colors.OKGREEN}Potensi Alamat IP v4 Ditemukan:{Colors.ENDC}")
            for ip in sorted(list(set(ip_addresses))): # Unik dan urutkan
                print(f"  -> {Colors.BOLD}{ip}{Colors.ENDC}")
                time.sleep(0.05)
        if urls:
            found = True
            print(f"\n  {Colors.OKGREEN}Potensi URL Ditemukan:{Colors.ENDC}")
            for url in sorted(list(set(urls))):
                # Batasi panjang URL yang ditampilkan jika terlalu panjang
                display_url = url[:80] + '...' if len(url) > 80 else url
                print(f"  -> {Colors.BOLD}{display_url}{Colors.ENDC}")
                time.sleep(0.05)
        if email_addresses:
            found = True
            print(f"\n  {Colors.OKGREEN}Potensi Alamat Email Ditemukan:{Colors.ENDC}")
            for email in sorted(list(set(email_addresses))):
                print(f"  -> {Colors.BOLD}{email}{Colors.ENDC}")
                time.sleep(0.05)

        if not found:
             print(f"\n  {Colors.GRAY}(Tidak ditemukan pola IP, URL, atau Email yang signifikan){Colors.ENDC}")

    except Exception as e:
        print(f"{Colors.FAIL}  [ERROR] Gagal menganalisis string: {e}{Colors.ENDC}")


def calculate_file_hash(file_path):
    """Menghitung dan menampilkan hash MD5 dan SHA256 dari file."""
    display_header("[PERHITUNGAN HASH FILE]")
    print(f"  Menghitung hash untuk: {Colors.BOLD}{file_path}{Colors.ENDC}\n")

    if not os.path.exists(file_path):
        print(f"{Colors.FAIL}[ERROR] File tidak ditemukan: {Colors.BOLD}{file_path}{Colors.ENDC}")
        return

    # MD5
    print(f"  {Colors.OKBLUE}--- Menghitung MD5 ---{Colors.ENDC}")
    md5_hash = calculate_md5(file_path) # Fungsi ini sudah punya progress bar & output
    if not md5_hash.startswith("[ERROR]"):
        print(f"  MD5 Hash\t: {Colors.BOLD}{md5_hash}{Colors.ENDC}")
    else:
        print(f"  MD5 Hash\t: {Colors.FAIL}{md5_hash}{Colors.ENDC}") # Tampilkan pesan error

    # SHA256
    print(f"\n  {Colors.OKBLUE}--- Menghitung SHA-256 ---{Colors.ENDC}")
    try:
        hasher = hashlib.sha256()
        with open(file_path, 'rb') as file:
            total_size = os.path.getsize(file_path)
            bytes_read = 0
            start_time = time.time()
            while chunk := file.read(8192): # Tingkatkan chunk
                hasher.update(chunk)
                bytes_read += len(chunk)
                elapsed_time = time.time() - start_time
                if total_size > 0 : # Hindari ZeroDivisionError
                    percent_complete = (bytes_read / total_size) * 100
                    if elapsed_time > 0:
                        speed = (bytes_read / 1024) / elapsed_time # Kecepatan dalam KB/s
                        print(f"\r{Colors.ORANGE}  Menghitung SHA-256: {percent_complete:.2f}% | {speed:.2f} KB/s", end="")
                    else:
                        print(f"\r{Colors.ORANGE}  Menghitung SHA-256: {percent_complete:.2f}%", end="")
                else:
                     print(f"\r{Colors.ORANGE}  Menghitung SHA-256...", end="")

        # Hapus baris progress setelah selesai
        print("\r" + " " * 70 + "\r", end="") # Bersihkan baris
        sha256_hash = hasher.hexdigest()
        print(f"{Colors.OKGREEN}  SHA-256 Hash selesai dihitung.{Colors.ENDC}")
        print(f"  SHA-256 Hash\t: {Colors.BOLD}{sha256_hash}{Colors.ENDC}")
    except FileNotFoundError: # Seharusnya sudah ditangani di awal
         print(f"{Colors.FAIL}  [ERROR] File tidak ditemukan: {Colors.BOLD}{file_path}{Colors.ENDC}")
    except PermissionError:
        print(f"{Colors.FAIL}  [ERROR] Tidak ada izin untuk membaca file: {Colors.BOLD}{file_path}{Colors.ENDC}")
    except Exception as e:
        print(f"{Colors.FAIL}  [ERROR] Kesalahan saat menghitung SHA-256: {e}{Colors.ENDC}")

    # Info Lokasi tidak relevan untuk file arbitrer
    print(f"\n  {Colors.GRAY}[INFO] Informasi lokasi GPS umumnya tidak tersedia untuk tipe file ini.{Colors.ENDC}")

def is_leap(year):
    """Cek apakah tahun kabisat."""
    return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)

def parse_nik(nik):
    """
    Fungsi untuk memvalidasi dan mengurai Nomor Induk Kependudukan (NIK) Indonesia.
    PERHATIAN: Membutuhkan data wilayah (provinsi, kab/kota, kecamatan) yang lengkap
               dan akurat di dictionary untuk hasil yang benar. Data saat ini terbatas.
    """
    display_header("[PARSE NIK]")
    print(f"  NIK yang diinput\t: {Colors.BOLD}{nik}{Colors.ENDC}")

    # 1. Validasi Dasar
    if not isinstance(nik, str) or not nik.isdigit() or len(nik) != 16:
        print(f"{Colors.FAIL}  [ERROR] Format NIK tidak valid. Harus 16 digit angka.{Colors.ENDC}")
        return None

    # 2. Parsing Komponen NIK
    try:
        kode_provinsi = nik[0:2]
        kode_kab_kota = nik[2:4]
        kode_kecamatan = nik[4:6]
        tanggal_lahir_raw = nik[6:8]
        bulan_lahir = nik[8:10]
        tahun_lahir_short = nik[10:12]
        no_urut = nik[12:] # Digit 13-16 adalah nomor urut unik

        # 3. Tentukan Jenis Kelamin dan Tanggal Lahir Sebenarnya
        tanggal_lahir_int = int(tanggal_lahir_raw)
        if tanggal_lahir_int > 40:
            jenis_kelamin = "Perempuan"
            tanggal_lahir_val = tanggal_lahir_int - 40
        else:
            jenis_kelamin = "Laki-laki"
            tanggal_lahir_val = tanggal_lahir_int

        # Format tanggal kembali ke string 2 digit dengan padding nol
        tanggal_lahir_str = f"{tanggal_lahir_val:02d}"

        # 4. Tentukan Tahun Lahir Lengkap (Heuristik)
        tahun_lahir_short_int = int(tahun_lahir_short)
        current_year = time.localtime().tm_year
        current_century_prefix = str(current_year)[:2] # e.g., '20'
        last_century_prefix = str(int(current_century_prefix) - 1) # e.g., '19'
        # Asumsi: Jika tahun lahir <= tahun sekarang (2 digit), gunakan abad ini (misal 20xx). Jika > tahun sekarang, gunakan abad lalu (misal 19xx)
        if tahun_lahir_short_int <= (current_year % 100):
            tahun_lahir_lengkap = int(f"{current_century_prefix}{tahun_lahir_short}")
        else:
            tahun_lahir_lengkap = int(f"{last_century_prefix}{tahun_lahir_short}")

        # 5. Validasi Bulan
        bulan_lahir_int = int(bulan_lahir)
        if not (1 <= bulan_lahir_int <= 12):
            print(f"{Colors.FAIL}  [ERROR] Bulan lahir ({bulan_lahir}) tidak valid.{Colors.ENDC}")
            return None

        # 6. Validasi Tanggal (menggunakan tahun lengkap dan tanggal yg sudah disesuaikan)
        max_tanggal = 31
        if bulan_lahir_int == 2:
            max_tanggal = 29 if is_leap(tahun_lahir_lengkap) else 28
        elif bulan_lahir_int in [4, 6, 9, 11]:
            max_tanggal = 30

        if not (1 <= tanggal_lahir_val <= max_tanggal):
             print(f"{Colors.FAIL}  [ERROR] Tanggal lahir ({tanggal_lahir_str} untuk bulan {bulan_lahir} tahun {tahun_lahir_lengkap}) tidak valid.{Colors.ENDC}")
             return None

        # 7. Dapatkan Nama Wilayah (dari dictionary terbatas)
        nama_provinsi = wilayah_dict.get(kode_provinsi, f"Tidak Diketahui ({kode_provinsi})")
        kode_prov_kab_kota = kode_provinsi + kode_kab_kota
        nama_kab_kota = kabupaten_kota_dict.get(kode_prov_kab_kota, f"Tidak Diketahui ({kode_kab_kota})")

        # Mencoba mencocokkan kode kecamatan (6 digit NIK) dengan kunci dictionary (7 digit asumsi)
        kecamatan_key_lookup = kode_prov_kab_kota + kode_kecamatan + '0'
        nama_kecamatan = kecamatan_dict.get(kecamatan_key_lookup, f"Tidak Diketahui ({kode_kecamatan})")
        if "Tidak Diketahui" in nama_kecamatan:
             print(f"{Colors.WARNING}  [INFO] Kode kecamatan {kode_kecamatan} (key: {kecamatan_key_lookup}) tidak ditemukan di data internal.{Colors.ENDC}")


        # 8. Tampilkan Hasil Parsing
        print(f"\n  {Colors.OKGREEN}--- Hasil Parsing NIK ---{Colors.ENDC}")
        print(f"  Provinsi\t\t: {Colors.BOLD}{nama_provinsi}{Colors.ENDC}")
        print(f"  Kabupaten/Kota\t: {Colors.BOLD}{nama_kab_kota}{Colors.ENDC}")
        print(f"  Kecamatan\t\t: {Colors.BOLD}{nama_kecamatan}{Colors.ENDC}")
        print(f"  Tanggal Lahir\t\t: {Colors.BOLD}{tanggal_lahir_str}-{bulan_lahir}-{tahun_lahir_lengkap}{Colors.ENDC}")
        print(f"  Jenis Kelamin\t\t: {Colors.BOLD}{jenis_kelamin}{Colors.ENDC}")
        print(f"  Nomor Urut\t\t: {Colors.BOLD}{no_urut}{Colors.ENDC}")

        return {
            "nik": nik,
            "provinsi": {"kode": kode_provinsi, "nama": nama_provinsi},
            "kabupaten_kota": {"kode": kode_kab_kota, "nama": nama_kab_kota},
            "kecamatan": {"kode": kode_kecamatan, "nama": nama_kecamatan},
            "tanggal_lahir": f"{tanggal_lahir_str}-{bulan_lahir}-{tahun_lahir_lengkap}",
            "jenis_kelamin": jenis_kelamin,
            "no_urut": no_urut
        }

    except ValueError:
        print(f"{Colors.FAIL}  [ERROR] NIK mengandung karakter non-angka setelah validasi awal.{Colors.ENDC}")
        return None
    except IndexError:
         print(f"{Colors.FAIL}  [ERROR] NIK terlalu pendek untuk diparsing (Index Error).{Colors.ENDC}")
         return None
    except Exception as e:
        print(f"{Colors.FAIL}  [ERROR] Gagal mengurai NIK: {e}{Colors.ENDC}")
        return None


def main_menu():
    """Fungsi utama yang menampilkan menu dan memproses pilihan pengguna."""
    while True:
        clear_screen()
        # Display header di sini agar muncul setiap kali menu tampil
        print("=" * 60)
        print(f"{Colors.HEADER}{Colors.BOLD}      ██████╗ ███████╗██╗  ██╗██████╗  ██████╗{Colors.ENDC}")
        print(f"{Colors.HEADER}{Colors.BOLD}      ██╔══██╗██╔════╝██║  ██║██╔══██╗██╔════╝{Colors.ENDC}")
        print(f"{Colors.HEADER}{Colors.BOLD}      ██████╔╝█████╗  ██║  ██║██████╔╝██║ ███╗{Colors.ENDC}")
        print(f"{Colors.HEADER}{Colors.BOLD}      ██╔══██╗██╔══╝   ╚██╗██╔╝██╔══██╗██║  ██║{Colors.ENDC}")
        print(f"{Colors.HEADER}{Colors.BOLD}      ██║  ██║███████╗  ╚████╔╝ ██║  ██║╚██████╔╝{Colors.ENDC}")
        print(f"{Colors.HEADER}{Colors.BOLD}      ╚═╝  ╚═╝╚══════╝   ╚═══╝  ╚═╝  ╚═╝ ╚═════╝ {Colors.ENDC}")
        print("=" * 60)
        print(f"{Colors.OKBLUE}          UBSIBER Forensic Tool - Owner: Satrio.U.N{Colors.ENDC}")
        print("=" * 60)
        print(f"{Colors.OKCYAN}{Colors.BOLD}                    MENU UTAMA{Colors.ENDC}")
        print("=" * 60)
        print(f"{Colors.OKGREEN}  [1] Analisis Gambar (Basic + Lokasi via Pillow){Colors.ENDC}")
        print(f"{Colors.OKGREEN}  [2] Analisis Gambar (Detail via ExifTool){Colors.ENDC}")
        print(f"{Colors.OKGREEN}  [3] Hash Password (SHA-256){Colors.ENDC}")
        print(f"{Colors.OKGREEN}  [4] Analisis Dokumen (Word .docx){Colors.ENDC}")
        print(f"{Colors.OKGREEN}  [5] Analisis String (Basic Hash){Colors.ENDC}")
        print(f"{Colors.OKGREEN}  [6] Analisis String (Cari IP, URL, Email){Colors.ENDC}")
        print(f"{Colors.OKGREEN}  [7] Hitung Hash File (MD5, SHA-256){Colors.ENDC}")
        print(f"{Colors.OKGREEN}  [8] Parse NIK Indonesia{Colors.ENDC}")
        print(f"{Colors.FAIL}  [0] Keluar{Colors.ENDC}")
        print("=" * 60)

        choice = input(f"{Colors.OKCYAN}Masukkan pilihan Anda [0-8]: {Colors.ENDC}").strip()
        wait_for_user = True # Defaultnya, tunggu user setelah aksi

        if choice == '1':
            image_path = input(f"{Colors.OKCYAN}Masukkan path gambar: {Colors.ENDC}").strip()
            clear_screen()
            analyze_image(image_path)
        elif choice == '2':
            image_path = input(f"{Colors.OKCYAN}Masukkan path gambar: {Colors.ENDC}").strip()
            clear_screen()
            analyze_image_exiftool(image_path)
        elif choice == '3':
            # Gunakan getpass untuk menyembunyikan input password
            import getpass
            try:
                 password = getpass.getpass(f"{Colors.OKCYAN}Masukkan password yang ingin di-hash (input tersembunyi): {Colors.ENDC}")
                 clear_screen()
                 hash_password(password)
            except Exception as e:
                 print(f"\n{Colors.FAIL}[ERROR] Tidak dapat menggunakan getpass: {e}{Colors.ENDC}")
                 # Fallback ke input biasa jika getpass gagal
                 password = input(f"{Colors.OKCYAN}Masukkan password yang ingin di-hash: {Colors.ENDC}")
                 clear_screen()
                 hash_password(password)

        elif choice == '4':
            doc_path = input(f"{Colors.OKCYAN}Masukkan path dokumen Word (.docx): {Colors.ENDC}").strip()
            clear_screen()
            analyze_document(doc_path)
        elif choice == '5':
            input_str = input(f"{Colors.OKCYAN}Masukkan string untuk dianalisis: {Colors.ENDC}")
            clear_screen()
            analyze_string(input_str)
        elif choice == '6':
            input_str = input(f"{Colors.OKCYAN}Masukkan string untuk dianalisis (lanjutan): {Colors.ENDC}")
            clear_screen()
            analyze_string_advanced(input_str)
        elif choice == '7':
             file_path = input(f"{Colors.OKCYAN}Masukkan path file untuk di-hash: {Colors.ENDC}").strip()
             clear_screen()
             calculate_file_hash(file_path)
        elif choice == '8':
             nik_input = input(f"{Colors.OKCYAN}Masukkan NIK (16 digit angka): {Colors.ENDC}").strip()
             clear_screen()
             parse_nik(nik_input)
        elif choice == '0':
            print(f"\n{Colors.VIOLET}Terima kasih telah menggunakan UBSIBER Forensic Tool. Keluar...{Colors.ENDC}")
            time.sleep(1)
            sys.exit() # Keluar dari program
        else:
            print(f"\n{Colors.FAIL}[ERROR] Pilihan tidak valid: '{choice}'. Silakan coba lagi.{Colors.ENDC}")
            wait_for_user = True # Tunggu sebentar agar user bisa baca error

        # Tunggu input pengguna sebelum kembali ke menu, kecuali saat keluar
        if wait_for_user:
             input(f"\n{Colors.GRAY}--- Tekan Enter untuk kembali ke menu ---{Colors.ENDC}")


if __name__ == "__main__":
    try:
        main_menu()
    except KeyboardInterrupt:
         print(f"\n\n{Colors.WARNING}Proses dihentikan oleh pengguna (Ctrl+C). Keluar.{Colors.ENDC}")
         sys.exit(1)
    except Exception as e:
         print(f"\n{Colors.FAIL}Terjadi error tak terduga di level utama: {e}{Colors.ENDC}")
         # Tampilkan traceback untuk debugging jika diperlukan
         # import traceback
         # traceback.print_exc()
         sys.exit(1)