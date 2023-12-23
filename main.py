from flask import Flask, render_template
import win32com.client

app = Flask(__name__)

def check_kms_status(server_url):
    try:
        # Membuat objek WMI
        wmi = win32com.client.GetObject("winmgmts:\\\\.\\root\\cimv2")

        # Menjalankan query untuk mendapatkan informasi KMS
        query = "SELECT * FROM SoftwareLicensingService"
        result = wmi.ExecQuery(query)

        # Memeriksa hasil query
        for item in result:
            if hasattr(item, "OA3xOriginalProductKey"):
                product_key = item.OA3xOriginalProductKey
                kms_host_name = item.DNSHostName
                kms_status = "Active" if item.GracefulShutdownStatus == 0 else "Not Active"
                return {"Product Key": product_key, "KMS Host Name": kms_host_name, "KMS Status": kms_status}

    except Exception as e:
        return {"Error": str(e)}

@app.route('/')
def index():
    # Daftar URL server KMS yang ingin diperiksa
    kms_servers = [
        "http://kms.iqbalrifai.eu.org",
        "http://kms.lux.iqbalrifai.eu.org",
        "https://kms.kor.iqbalrifai.eu.org",
        "https://kms.hk.iqbalrifai.eu.org"
        # Tambahkan URL server KMS lainnya sesuai kebutuhan
    ]

    results = []
    for server in kms_servers:
        kms_status = check_kms_status(server)
        results.append({"Server": server, "Status": kms_status})

    return render_template('index.html', results=results)

if __name__ == "__main__":
    app.run(debug=True)