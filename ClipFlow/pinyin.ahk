; #Include "%A_ScriptDir%\dict.ahk"
#Include "../dict.ahk"

regionISOMap := Map()
loop regionISO.Length {
    regionISOMap[regionISO[A_Index][1]] := regionISO[A_Index][2]
}

provinceMap := Map()
loop province.Length {
    provinceMap[province[A_Index][1]] := province[A_Index][2]
}

dictionaryMap := Map()
loop dictionary.Length {
    ; key: han, value: pinyin
    dictionaryMap[dictionary[A_Index][2]] := dictionary[A_Index][1]
}

doubleLastNameMap := Map()
loop doubleLastName.Length {
    ; key: han, value: pinyin
    doubleLastNameMap[doubleLastName[A_Index][1]] := doubleLastName[A_Index][2]
}

getCountryCode(country) {
    for region, code in regionISOMap {
        if (country = region) {
            return code
        }
    }
}

getProvince(address) {
    for province, code in provinceMap {
        if (InStr(address, province)) {
            return code
        }
    }
}

getPinyin(hanzi) {
    for han, pinyin in dictionaryMap {
        if (InStr(han, hanzi)) {
            return pinyin
        }
    }
}

getFullnamePinyin(fullname) {
    ; is a double hanzi last name
    for han, pinyin in doubleLastNameMap {
        if (SubStr(fullname, 1, 2) = han) {
            ; return [lastname, name]
            return [SubStr(fullname, 1, 2), SubStr(fullname, 3)]
        }
    }
    ; one hanzi last name
    lastname := getPinyin(SubStr(fullname, 1, 1))
    firstnameSplit := StrSplit(SubStr(fullname, 2), "")
    firstname := ""
    loop firstnameSplit.Length {
        firstname .= getPinyin(firstnameSplit[A_Index]) . " "
    }
    return [lastname, firstname]
}