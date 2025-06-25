def number_to_ordinals(number_str):
    match number_str:
        case "1" | "3":
            number_str += r"\textsuperscript{er}"
        case "2":
            number_str += r"\textsuperscript{do}"
        case "4" | "5" | "6":
            number_str += r"\textsuperscript{to}"
        case "7" | "10":
            number_str += r"\textsuperscript{mo}"
        case "8":
            number_str += r"\textsuperscript{vo}"
        case "9":
            number_str += r"\textsuperscript{no}"
    return number_str 