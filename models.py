class Reaction:
    def __init__(self):
        self.emoji = ""
        self.time = ""
        self.orgid = ""

    def toString(self):
        return "emoji: {0} || Time: {1} || Orgid: {2}".format(self.emoji, self.time, self.orgid)


class MensagemCompleta:
    def __init__(self):
        self.message = ""
        self.sender = ""
        self.time = ""
        self.files = []
        self.hasEmoji = False
        self.isMention = False
        self.mention = ""
        self.reactions = []

    def toString(self):
        return "Message: {0} || Time: {1} || Sender: {2} || IsMention: {3} || Mention: {4} || HasEmoji: {5}".format(self.message,
                                                                                                   self.time,
                                                                                                   self.sender,
                                                                                                   str(self.isMention),
                                                                                                   self.mention,str(self.hasEmoji))


class Contacto:
    def __init__(self, nome, email, orgid):
        self.nome = nome
        self.email = email
        self.orgid = orgid

    def toString(self):
        return "Nome: {0} || Email: {1} || Orgid {2}".format(self.nome, self.email, self.orgid)


class Chamada:
    def __init__(self, criador, timestart, timefinish):
        self.criador = criador
        self.timestart = timestart
        self.timefinish = timefinish
        self.presentes = []
