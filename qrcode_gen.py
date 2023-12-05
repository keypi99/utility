import qrcode
import base64
import io


def generate_qrcode(url):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )

    qr.add_data(url)
    qr.make(fit=True)
    path=str("C:/Users/polik/Downloads/name.png")
    img_qrcode = qr.make_image(fill_color="black", back_color="white")
    temp = io.BytesIO()
    img_qrcode.save(path, format="PNG")
    qr_img = base64.b64encode(temp.getvalue())

if __name__ == "__main__":
    generate_qrcode(url='https://www.jw.org/finder?wtlocale=I&docid=502016853&srcid=share')
