import pygame

class fontPics():
    def __init__(self):
        pygame.init()
        screen = pygame.display.set_mode((100,50))
        screen.fill((255,255,255))
        self.getFonts(screen)
        pygame.quit()

    def getFonts(self,screen):
        fnts = pygame.font.get_fonts()
        self.renderFonts(fnts,screen)
    
    def renderFonts(self,fonts,screen):
        for fnt in fonts:
            self.typeOnScreen(fnt,screen)

    def typeOnScreen(self,fontName,screen):
        fnt = pygame.font.SysFont(fontName, 20)
        text = "AaBbYyZz"
        textPic = fnt.render(text,1,(0,0,0))
        screen.blit(textPic,(3,3))
        self.makeImage(screen,fontName)
     
    def makeImage(self,screen,fontName):
        pygame.image.save(screen,str(fontName+".BMP"))
        screen.fill((255,255,255))

if __name__ == "__main__":
    fontPics()