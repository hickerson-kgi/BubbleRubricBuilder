#!/usr/bin/env python3
'''
Minimal Kivy program example
'''
import kivy
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout



class Main(BoxLayout):
    def tell_me(self, value):
        if self.ids['text_id'].focus:
            print('Just Focused!', value)
        else:
            print('Unfocused', value)

    def tab_test(self, text):
        # all_ids = self.ids
        # for id in all_ids:
        #     print(id)
        print(text)

class IntroKivyApp(App):
    def build(self):
        return Main()

IntroKivyApp().run()
