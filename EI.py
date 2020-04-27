import codecs
import os
from pathlib import Path
from models import Contacto, MensagemCompleta, Reaction

appdata = os.getenv('APPDATA')
home = str(Path.home())
levelDBPath = appdata + "\\Microsoft\\Teams\\IndexedDB\\https_teams.microsoft.com_0.indexeddb.leveldb"
levelDBPathLost = appdata + "\\Microsoft\\Teams\\IndexedDB\\https_teams.microsoft.com_0.indexeddb.leveldb\\lost"
projetoEIAppDataPath = appdata + "\\ProjetoEI\\"
logFinal = open(os.path.join(projetoEIAppDataPath, "logTotal.txt"), "a+", encoding="utf-8")
logMsgs = open(os.path.join(projetoEIAppDataPath, "logMsgs.txt"), "a+", encoding="utf-8")
logCall = open(os.path.join(projetoEIAppDataPath, "logCalls.txt"), "a+", encoding="utf-8")
linkEmoji = "https://statics.teams.cdn.office.net/evergreen-assets/skype/v2/{0}/50.png"
arrayMensagens = []
arrayContactos = []


def find(pt, path):
    fls = []
    for root, dirs, files in os.walk(path):
        for name in files:
            if pt in name:
                fls.append(os.path.join(root, name))
    return fls


def acentuar(acentos, st):
    for hexa, acento in acentos.items():
        st = st.replace(hexa, acento)
    return st


def decoder(string, pattern):
    s = string.encode("utf-8").find(pattern.encode("utf-16le"))
    l = string[s - 1:]
    t = l.encode("utf-16le")
    st = str(t)
    st = st.replace(r"\x00", "")
    st = st.replace(r"\x", "feff00")
    st = st.replace(r"b'", "")
    return st


def multiFind(text):
    f = ""
    index = 0
    a = ""
    l = list(text)
    acentos = {}
    while index < len(text):
        index = text.find('feff00', index)
        if index == -1:
            break
        for x in range(index, index + 8):
            f = f + l[x]
        a = bytes.fromhex(f).decode("utf-16be")
        acentos[f] = a
        index += 8  # +6 because len('feff00') == 2
        f = ""
    return acentos


def utf16customdecoder(string, pattern):
    st = decoder(string, pattern)
    acentos = multiFind(st)
    st = acentuar(acentos, st)
    return st


def testeldb(path):
    from subprocess import Popen, PIPE
    import ast
    for f in os.listdir(path):
        if f.endswith(".ldb") and os.path.getsize(path + "\\" + f) > 0:
            process = Popen([os.path.join(home, r"PycharmProjects\EI\ldbdump"), os.path.join(path, f)], stdout=PIPE)
            (output, err) = process.communicate()
            try:
                result = output.decode("utf-8")
                for line in (result.split("\n")[1:]):
                    if line.strip() == "": continue
                    parsed = ast.literal_eval("{" + line + "}")
                    key = list(parsed.keys())[0]
                    logFinal.write(parsed[key])
            except:
                print("falhou")


def writelog(path):
    log = open(path, "r", encoding="mbcs")
    while line := log.readline():
        logFinal.write(line)
    log.close()


def crialogtotal():
    os.system(r'cmd /c "LDBReader\PrjTeam.exe"')
    testeldb(levelDBPathLost)
    logLost = find(".log", levelDBPathLost)
    writelog(os.path.join(projetoEIAppDataPath, "logIndexedDB.txt"))
    logsIndex = find(".log", levelDBPath)
    for path in logLost:
        writelog(path)
    for path in logsIndex:
        writelog(path)
    logFinal.close()


def filtro(buffer):
    filtros = {
        "trimmedMessageContent": 0,
        "RichText/Html": 0,
        "renderContent": 0,
        "skypeguid": 0
    }
    sender = ""
    time = ""
    indexTexto = 0
    indexTextoFinal = 0
    name = ""
    isEmoji = False
    isRichHtml = False
    isRichHtmlDone = False
    richHtml = ""
    isReaction = False
    reaction = Reaction()
    emojiReaction = ""
    orgidReaction = ""
    timeReaction = ""
    message = ""
    reactions = []
    ok = 1
    nAspas = False
    for filter in filtros:
        for line in buffer:
            if filter in line:
                filtros[filter] = 1
    for x in filtros.values():
        if x == 0:
            ok = 0

    if ok:
        logMsgs.write("-------------------------------------------------------------------------------\n")
        isMention = False
        mention = ""
        for line in buffer:
            logMsgs.write(line)
            name = ""
            if "imdisplayname" in line:
                indexTexto = line.find("imdisplayname")
                indexTextoFinal = line.find("skypeguid")
                l = list(line)
                for x in range(indexTexto, indexTextoFinal):
                    if "\"" in l[x]:
                        nAspas = not nAspas
                        continue
                    if nAspas:
                        name = name + l[x + 1]
                name = name.replace("\"", "")
                sender = name
                nAspas = False
            name = ""

            if "composetime" in line:
                indexTextoFinal = line.find("type")
                l = list(line)
                for x in range(0, indexTextoFinal - 1):
                    name = name + l[x]
                name = name.replace("\"", "")
                time = name
                nAspas = False

            if "RichText/Html" in line:
                isRichHtml = True
                richHtml += line

            if "renderContent" in line:
                isRichHtml = False
                isRichHtmlDone = True

            if isRichHtml:
                richHtml += line

            if isRichHtmlDone:
                if "<div>".encode("utf-16le") in line.encode("utf-8"):
                    st = utf16customdecoder(richHtml, "<div>")
                    richHtml = st
                    if "https://statics.teams.cdn.office.net/evergreen-assets/skype/v2/" in st:
                        # Não pronto para multiplos emojis
                        indexMultiplosEmojis = 0
                        indexSpan = 0
                        while indexSpan != -1:

                            indexSpan = st.find("<span class=\"animated-emoticon", indexMultiplosEmojis)
                            indexMultiplosEmojis = indexSpan
                            indexSpanfinal = st.find("</span>", indexSpan)
                            indexSpanfinal += 7
                            l = list(st)
                            span = ""
                            emoji = ""
                            for x in range(indexSpan, indexSpanfinal):
                                span = span + l[x]

                            indexEmoji = span.find("itemid=\"")
                            indexEmojiFinal = span.find(" itemscope")
                            indexEmoji += 7
                            spanList = list(span)
                            for x in range(indexEmoji + 1, indexEmojiFinal - 1):
                                emoji += spanList[x]
                            st = st.replace(span, " " + linkEmoji.format(emoji) + " ")
                            st = st.replace(r"\n", "")
                            richHtml = st
                if "http://schema.skype.com/Mention" in richHtml:
                    indexMencionado = 0
                    span = ""
                    while indexMencionado != -1:
                        isMention = True
                        indexMencionado = richHtml.find("<span itemscope itemtype=\"http://schema.skype.com/Mention\"",
                                                        indexMencionado)
                        indexMencionadoFinal = richHtml.find("</span>", indexMencionado)
                        l = list(richHtml)
                        for x in range(indexMencionado, indexMencionadoFinal):
                            span += l[x]

                        indexSpanMention = span.find("itemid") + 11
                        s = list(span)
                        for x in range(indexSpanMention, len(s)):
                            name += s[x]

                        richHtml = richHtml.replace(span + "</span>", "Mention: {0}".format(name))

                        span = ""
                        mention = name
                        name = ""

                indexTexto = richHtml.find("<div")
                indexTexto += 5
                indexTextoFinal = richHtml.find("</div>")
                l = list(richHtml)
                for x in range(indexTexto, indexTextoFinal):
                    name = name + l[x]
                name = name.replace("<div>", "")
                message = name
                # if message.find(
                #         "https://statics.teams.cdn.office.net/evergreen-assets/skype/v2/") != -1 or message.find(
                #     "Mention: ") != -1:
                #     print(message)
                #     print("---------------")
                nAspas = False
                name = ""

                isRichHtmlDone = False

            if "mri" in line and "orgid" in line and ("creatorProfile" not in line and "mention" not in line):
                indexOrgid = line.find("orgid")
                indexOrgid += 5
                indexOrgidFinal = line.find("timeN")
                indexOrgidFinal -= 2
                indexKey = line.find("key\"")
                indexKey += 5
                indexKeyFinal = line.find("user")
                indexKeyFinal -= 2
                l = list(line)
                for x in range(indexKey, indexKeyFinal):
                    name = name + l[x]
                if name != "":
                    emojiReaction = name
                    name = ""
                for x in range(0, 13):
                    name = name + l[x]
                timeReaction = name
                name = ""
                for x in range(indexOrgid, indexOrgidFinal):
                    name = name + l[x]
                orgidReaction = name
                name = ""
                reaction.time = timeReaction
                reaction.emoji = emojiReaction
                reaction.orgid = orgidReaction
                reactions.append(reaction)

        logMsgs.write("-------------------------------------------------------------------------------\n")
        isEmoji = False
        # for r in reactions:
        #     print(r.toString())
        # print("--------------------------------------------------------------------------------")
        mensagem = MensagemCompleta()
        if isMention:
            mensagem.isMention = True
            mensagem.mention = mention
        mensagem.message = message
        mensagem.time = time
        mensagem.sender = sender
        if isReaction:
            mensagem.reactions = reactions
        if isEmoji:
            mensagem.hasEmoji = True
        arrayMensagens.append(mensagem)


def filtroReverso(buffer):
    filtros = {
        "notifications": 1,
        "skypeMessageLocal": 1,
        "messageStorageStateI": 1
    }
    ok = 1
    for filter in filtros:
        for line in buffer:
            if filter in line:
                filtros[filter] = 0

    for x in filtros.values():
        if x == 0:
            ok = 0

    if ok:
        logCall.write("-------------------------------------------------------------------------------\n")
        for line in buffer:
            logCall.write(line)
        logCall.write("-------------------------------------------------------------------------------\n")


def findpadrao():
    logFinalRead = open(os.path.join(projetoEIAppDataPath, "logTotal.txt"), "r", encoding="utf-8")
    flagMgs = 0
    buffer = []
    while line := logFinalRead.readline():
        if flagMgs == 1:
            buffer.append(line)
        if "RichText/Html" in line:
            flagMgs = 1
            buffer.append(line)
        if "pinnedTime_" in line:
            flagMgs = 0
            filtro(buffer)
            buffer.clear()
    logFinalRead.close()
    for m in arrayMensagens:
        print(m.toString())
    logFinalRead = open(os.path.join(projetoEIAppDataPath, "logTotal.txt"), "r", encoding="utf-8")
    while line := logFinalRead.readline():

        if flagMgs == 1:
            buffer.append(line)
        if "conversationId" in line:
            flagMgs = 1
            buffer.append(line)
        if "parentMessageId\"" in line:
            flagMgs = 0
            filtroReverso(buffer)
            buffer.clear()
    logMsgs.close()
    logCall.close()


def geraContactos():
    logFinalRead = open(os.path.join(projetoEIAppDataPath, "logTotal.txt"), "r", encoding="utf-8")
    orgid = ""
    name = ""
    mail = ""
    nAspas = False
    while line := logFinalRead.readline():
        if "itemRank" in line:
            indexNome = line.find("displayName")
            indexNomeFinal = line.find("userPrincipalName")
            indexMail = line.find("email")
            indexMailFinal = line.find("description")
            indexObject = line.find("objectId")
            indexObjectEnd = line.find("$$lastname_lowercase")
            lista = list(line)

            for a in range(indexNome, indexNomeFinal):
                if "\"" in lista[a]:
                    nAspas = not nAspas
                    continue
                if nAspas:
                    name = name + lista[a + 1]
            name = name.replace("\"", "")
            nameContacto = name
            nAspas = False

            for x in range(indexMail, indexMailFinal):
                if "\"" in lista[x]:
                    nAspas = not nAspas
                    continue
                if nAspas:
                    mail = mail + lista[x + 1]
            mail = mail.replace("\"", "")
            emailContacto = mail
            nAspas = False

            for z in range(indexObject, indexObjectEnd):
                if "\"" in lista[z]:
                    nAspas = not nAspas
                    continue
                if nAspas:
                    orgid = orgid + lista[z + 1]

            orgid = orgid.replace("\"", "")
            orgidContacto = orgid

            orgid = ""
            name = ""
            mail = ""
            nAspas = False

            contacto = Contacto(nameContacto, emailContacto, orgidContacto)
            arrayContactos.append(contacto)


if __name__ == "__main__":
    crialogtotal()
    findpadrao()
    geraContactos()
