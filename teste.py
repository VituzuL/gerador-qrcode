import cv2
import numpy as np
from kivy.lang import Builder
from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivy.uix.image import Image
from kivy.clock import Clock

KV = '''
MDScreen:
    MDBoxLayout:
        orientation: 'vertical'

        MDToolbar:
            title: 'Leitor de QR Code'
            left_action_items: [['menu', lambda x: app.menu()]]

        Image:
            id: img_camera
            size_hint: (1, 0.8)

        MDLabel:
            id: label_result
            text: 'Resultado do QR Code aparecerá aqui.'
            halign: 'center'

        MDBoxLayout:
            orientation: 'horizontal'
            size_hint: (1, 0.1)

            MDRectangleFlatButton:
                text: 'Iniciar Leitura'
                on_release: app.start_reading()

            MDRectangleFlatButton:
                text: 'Salvar'
                on_release: app.save_data()
'''

class QRCodeApp(MDApp):
    def build(self):
        return Builder.load_string(KV)

    def start_reading(self):
        self.capture = cv2.VideoCapture(0)  # 0 para câmera padrão
        Clock.schedule_interval(self.update_frame, 1.0 / 30.0)  # 30 fps

    def get_texture(self, buf):
        # Converte a imagem do OpenCV para o formato que o Kivy pode usar
        buf = buf.tostring()
        texture = Image(size=(buf.shape[1], buf.shape[0]))
        texture.texture.blit_buffer(buf, colorfmt='rgb', bufferfmt='ubyte')
        return texture.texture

    def display_result(self, result):
        self.root.ids.label_result.text = f'Resultado: {result}'

    def save_data(self):
        result = self.root.ids.label_result.text
        print(f'Dados salvos: {result}')  # Aqui você pode salvar em um arquivo ou em outro lugar.

    def on_stop(self):
        if hasattr(self, 'capture'):
            self.capture.release()

if __name__ == '__main__':
    QRCodeApp().run()
