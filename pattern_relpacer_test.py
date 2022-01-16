
import pattern_relpacer

class TestWordPatternReplacer:
    def setup(self):
        self.__wp = pattern_relpacer.WordPatternReplacer()

    def test_read_word(self):
        self.__wp.read_word("template.docx")
        

