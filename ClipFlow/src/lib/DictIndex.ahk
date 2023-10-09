; #Include "%A_ScriptDir%\src\lib\Dict.ahk"
#Include "../Dict.ahk"

regionISOMap := Map()
loop regionISO.Length {
    ; region : code
    regionISOMap[regionISO[A_Index][1]] := regionISO[A_Index][2]
}

provinceMap := Map()
loop province.Length {
    ; province name : code
    provinceMap[province[A_Index][1]] := province[A_Index][2]
}

dictionaryMap := Map()
loop dictionary.Length {
    ; hanzi : pinyin
    dictionaryMap[dictionary[A_Index][2]] := dictionary[A_Index][1]
}

doubleLastNameMap := Map()
loop doubleLastName.Length {
    ; hanzi : pinyin
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
    lastname := (doubleLastName.Has(SubStr(fullname, 1, 2)))
        ? doubleLastName[SubStr(fullname, 1, 2)]
        : getPinyin(SubStr(fullname, 1, 1))

    firstnameSplit := StrSplit(SubStr(fullname, 2), "")
    firstname := ""
    loop firstnameSplit.Length {
        firstname .= getPinyin(firstnameSplit[A_Index]) . " "
    }
    return [lastname, firstname]
}