"""Lokalni prijevodni sloj za BTG similarity.

Ovdje nisu Amazon node ID mapiranja, nego sigurnosni rječnik pojmova koji pomaže
usporediti DE kategoriju s NL/PL/IE/SE/FR/IT/ES BTG putanjama bez vanjskog AI-ja.

Kad vidiš da neki pojam često fali, dodaj ga u postojeći concept ili napravi novi.
Ključ concepta mora biti stabilan engleski naziv, a vrijednosti su sinonimi po jezicima.
"""

CONCEPT_TRANSLATIONS = {
    "storage": [
        "aufbewahren", "aufbewahrung", "ordnen", "ordnung", "organisieren", "organizer", "organizer",
        "storage", "storing", "organising", "organizing", "organisation", "organization",
        "rangement", "organisation", "organisateur", "guardar", "almacenamiento", "ordenacion", "ordenación",
        "conservazione", "organizzazione", "contenitori", "contenitore",
        "opbergen", "opslag", "ordenen", "organiseren", "bewaren",
        "przechowywanie", "organizacja", "organizery", "pojemniki", "pojemnik",
        "forvaring", "förvaring", "organisering", "förvarings", "opbevaring",
    ],
    "waste": [
        "abfall", "müll", "muell", "mull", "recycling", "mülltonne", "mülltonnen", "abfalleimer", "papierkorb",
        "trash", "waste", "garbage", "rubbish", "recycle", "recycling", "bin", "bins", "dustbin", "trashcan", "trash cans",
        "déchets", "dechets", "ordures", "poubelle", "poubelles", "recyclage",
        "basura", "residuos", "reciclaje", "cubo de basura", "papelera", "contenedor de basura",
        "rifiuti", "spazzatura", "riciclaggio", "pattumiera", "pattumiere", "cestino",
        "afval", "recycling", "prullenbak", "prullenbakken", "vuilnisbak", "vuilnisbakken",
        "śmieci", "smieci", "odpady", "recykling", "kosz", "kosze", "kosze na śmieci", "kosze na smieci", "kubły", "kubly",
        "avfall", "sopor", "återvinning", "atervinning", "soptunna", "soptunnor", "papperskorg", "skräp", "skrap",
    ],
    "garden": [
        "garten", "gartenarbeit", "rasen", "outdoor", "aussen", "außen",
        "garden", "gardening", "lawn", "yard", "outdoor",
        "jardin", "jardinage", "extérieur", "exterieur", "riego", "arrosage",
        "jardín", "jardin", "jardinería", "jardineria", "exterior",
        "giardino", "giardinaggio", "esterno",
        "tuin", "tuinieren", "buiten",
        "ogród", "ogrod", "ogrodowe", "ogrodnictwo", "ogród i taras", "ogrod i taras",
        "trädgård", "tradgard", "trädgårds", "tradgards", "utomhus",
    ],
    "watering": [
        "bewässerung", "bewasserung", "schlauch", "schlauchsysteme", "wasser", "bewässern",
        "watering", "irrigation", "hose", "hoses", "water", "water supply", "sprinkler",
        "arrosage", "irrigation", "tuyau", "eau", "tuyaux", "asperseur",
        "riego", "manguera", "agua", "aspersor", "irrigación", "irrigacion",
        "irrigazione", "tubo", "acqua", "annaffiare", "irrigatore",
        "bewatering", "irrigatie", "tuinslang", "water", "sproeier",
        "nawadniania", "nawadnianie", "wąż", "waz", "węże", "weze", "woda", "wodę", "wode", "zraszacz",
        "bevattning", "slang", "vatten", "vattning", "sprinkler",
    ],
    "tank": [
        "wassertank", "wassertanks", "tank", "tanks", "zisterne", "zisternen", "unterirdische wassertanks",
        "water tank", "water tanks", "tank", "tanks", "cistern", "cisterns", "reservoir",
        "réservoir", "reservoir", "réservoirs", "reservoirs", "citerne", "citernes",
        "depósito", "deposito", "depósitos", "depositos", "cisterna", "cisternas", "tanque", "tanques",
        "serbatoio", "serbatoi", "cisterna", "cisterne", "tanica", "taniche",
        "watertank", "watertanks", "waterreservoir", "regenton", "reservoir",
        "zbiornik", "zbiorniki", "zbiornik na wodę", "zbiorniki na wodę", "zbiornik na wode", "zbiorniki na wode", "cysterna", "cysterny",
        "vattentank", "vattentankar", "cistern", "vattencistern", "vattenbehållare", "vattenbehallare",
    ],
    "plant_pots": [
        "blumentopf", "blumentöpfe", "blumentopfe", "pflanzgefäß", "pflanzgefäss", "pflanzgefass", "pflanzgefäße", "pflanzgefaesse", "pflanzkübel", "pflanzkubel",
        "plant pot", "plant pots", "planter", "planters", "flowerpot", "flowerpots", "flower pot", "flower pots",
        "pot de fleur", "pots de fleurs", "jardinière", "jardiniere", "bac à fleurs", "bac a fleurs",
        "maceta", "macetas", "jardinera", "jardineras", "tiesto", "tiestos",
        "vaso", "vasi", "fioriera", "fioriere", "portavasi",
        "bloempot", "bloempotten", "plantenbak", "plantenbakken", "potten",
        "donica", "donice", "doniczki", "osłonka", "oslonka", "osłonki", "oslonki", "kwietnik", "kwietniki",
        "kruka", "krukor", "blomkruka", "blomkrukor", "planteringskärl", "planteringskarl",
    ],
    "kitchen": [
        "küche", "kuche", "kueche", "küchen", "haushalt", "haus und küche",
        "kitchen", "home kitchen", "cooking", "cook", "home and kitchen",
        "cuisine", "maison et cuisine", "cocina", "hogar y cocina", "cucina", "casa e cucina",
        "keuken", "huis en keuken", "kuchnia", "dom i kuchnia", "kök", "kok", "hem och kök",
    ],
    "cookware": [
        "kochgeschirr", "kochen", "pfanne", "pfannen", "topf", "töpfe", "toepfe", "küchenutensilien", "kuchenutensilien",
        "cookware", "utensil", "utensils", "pots", "pans", "kitchenware",
        "ustensiles", "batterie de cuisine", "casseroles", "poêles", "poeles",
        "utensilios", "menaje", "ollas", "sartenes", "batería de cocina", "bateria de cocina",
        "utensili", "pentole", "padelle", "batteria da cucina",
        "kookgerei", "keukengerei", "pannen", "potten",
        "naczynia kuchenne", "garnki", "patelnie", "przybory kuchenne", "akcesoria kuchenne",
        "köksredskap", "koksredskap", "kastruller", "stekpannor", "köksutrustning", "koksutrustning",
    ],
    "cleaning": [
        "reinigen", "reinigung", "putzen", "sauberkeit", "cleaning", "clean", "nettoyage", "limpieza", "pulizia",
        "schoonmaak", "czyszczenie", "sprzątanie", "sprzatanie", "städning", "stadning",
    ],
    "tools": [
        "werkzeug", "werkzeuge", "tools", "tool", "outils", "herramientas", "attrezzi", "gereedschap",
        "narzędzia", "narzedzia", "verktyg",
    ],
    "automotive": [
        "automotive", "auto", "autos", "car", "cars", "fahrzeug", "fahrzeuge", "samochód", "samochod", "samochody", "bil", "bilar", "voiture", "coche", "autoaccessoires",
    ],
}

# Ako se concept pojavi u zadnjem dijelu DE putanje, target mora imati isti concept.
# Time se izbjegava npr. Wassertanks -> Donice ili Mülltonnen -> Naczynia kuchenne.
CRITICAL_LEAF_CONCEPTS = {"waste", "tank", "plant_pots", "cookware", "automotive"}

# Parovi koji se ne smiju prihvatiti samo zato što imaju jednu opću riječ zajedničku.
OPPOSING_CONCEPTS = {
    "waste": {"cookware", "plant_pots"},
    "tank": {"plant_pots", "cookware", "waste"},
    "plant_pots": {"tank", "waste", "cookware"},
    "cookware": {"waste", "tank", "plant_pots"},
    "kitchen": {"garden"},
}
