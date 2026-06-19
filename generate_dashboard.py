#!/usr/bin/env python3
"""
Generador del dashboard de la Porra Mundial 2026 (Reveni).

Lee el Excel de administración (Porra_Admin_v4_EN.xlsx) y produce un único
fichero HTML autocontenido con analíticas estilo "quién es el más rebelde",
"quién piensa como quién", estilo de apuesta, favoritos, etc.

Reejecutable: cuando se rellenen resultados/eliminatorias en el Excel, basta
con volver a correr el script y el HTML se regenera (incluyendo aciertos en
directo si ya hay resultados).

Uso:
    python3 generate_dashboard.py [ruta_xlsx] [ruta_salida_html]
"""

from __future__ import annotations

import json
import sys
import statistics
from collections import Counter, defaultdict
from datetime import date as _date
from pathlib import Path

import openpyxl

# --------------------------------------------------------------------------
# Configuración / constantes
# --------------------------------------------------------------------------

DEFAULT_XLSX = "Porra_Admin_v4_EN.xlsx"
DEFAULT_OUT = "index.html"  # raíz del sitio de GitHub Pages
LOGO_SVG_PATH = "reveni-logo.svg"

N_PARTICIPANTS_MAX = 30          # huecos P1..P30 en el Excel
RAW_FIRST_HOME_COL = 3           # columna C = primer "home" de P1
NAME_ROW = 6                     # fila con los nombres reales
GROUP_MATCH_ROWS = range(7, 79)  # 72 partidos de fase de grupos
QUALIFIER_ROWS = range(80, 116)  # 1st/2nd/3rd de cada grupo (A..L)
THIRDS_ROWS = range(117, 125)    # 8 mejores terceros

MATCH_DATES = {
    "GA-M1": "2026-06-11", "GA-M2": "2026-06-11",
    "GA-M3": "2026-06-18", "GA-M4": "2026-06-18",
    "GA-M5": "2026-06-24", "GA-M6": "2026-06-24",
    "GB-M1": "2026-06-12", "GB-M2": "2026-06-13",
    "GB-M3": "2026-06-18", "GB-M4": "2026-06-18",
    "GB-M5": "2026-06-24", "GB-M6": "2026-06-24",
    "GC-M1": "2026-06-13", "GC-M2": "2026-06-13",
    "GC-M3": "2026-06-19", "GC-M4": "2026-06-19",
    "GC-M5": "2026-06-24", "GC-M6": "2026-06-24",
    "GD-M1": "2026-06-12", "GD-M2": "2026-06-13",
    "GD-M3": "2026-06-19", "GD-M4": "2026-06-19",
    "GD-M5": "2026-06-25", "GD-M6": "2026-06-25",
    "GE-M1": "2026-06-14", "GE-M2": "2026-06-14",
    "GE-M3": "2026-06-20", "GE-M4": "2026-06-20",
    "GE-M5": "2026-06-25", "GE-M6": "2026-06-25",
    "GF-M1": "2026-06-14", "GF-M2": "2026-06-14",
    "GF-M3": "2026-06-20", "GF-M4": "2026-06-20",
    "GF-M5": "2026-06-25", "GF-M6": "2026-06-25",
    "GG-M1": "2026-06-15", "GG-M2": "2026-06-15",
    "GG-M3": "2026-06-21", "GG-M4": "2026-06-21",
    "GG-M5": "2026-06-26", "GG-M6": "2026-06-26",
    "GH-M1": "2026-06-15", "GH-M2": "2026-06-15",
    "GH-M3": "2026-06-21", "GH-M4": "2026-06-21",
    "GH-M5": "2026-06-26", "GH-M6": "2026-06-26",
    "GI-M1": "2026-06-16", "GI-M2": "2026-06-16",
    "GI-M3": "2026-06-22", "GI-M4": "2026-06-22",
    "GI-M5": "2026-06-26", "GI-M6": "2026-06-26",
    "GJ-M1": "2026-06-16", "GJ-M2": "2026-06-16",
    "GJ-M3": "2026-06-22", "GJ-M4": "2026-06-22",
    "GJ-M5": "2026-06-27", "GJ-M6": "2026-06-27",
    "GK-M1": "2026-06-17", "GK-M2": "2026-06-17",
    "GK-M3": "2026-06-23", "GK-M4": "2026-06-23",
    "GK-M5": "2026-06-27", "GK-M6": "2026-06-27",
    "GL-M1": "2026-06-17", "GL-M2": "2026-06-17",
    "GL-M3": "2026-06-23", "GL-M4": "2026-06-23",
    "GL-M5": "2026-06-27", "GL-M6": "2026-06-27",
}

# Traducción EN -> ES (reconocible) + bandera
TEAMS = {
    "Mexico": ("México", "🇲🇽"), "South Africa": ("Sudáfrica", "🇿🇦"),
    "South Korea": ("Corea del Sur", "🇰🇷"), "Czech Rep.": ("Rep. Checa", "🇨🇿"),
    "Canada": ("Canadá", "🇨🇦"), "Bosnia-Herz.": ("Bosnia", "🇧🇦"),
    "Qatar": ("Catar", "🇶🇦"), "Switzerland": ("Suiza", "🇨🇭"),
    "Brazil": ("Brasil", "🇧🇷"), "Morocco": ("Marruecos", "🇲🇦"),
    "Haiti": ("Haití", "🇭🇹"), "Scotland": ("Escocia", "🏴󠁧󠁢󠁳󠁣󠁴󠁿"),
    "USA": ("EE. UU.", "🇺🇸"), "Paraguay": ("Paraguay", "🇵🇾"),
    "Australia": ("Australia", "🇦🇺"), "Turkey": ("Turquía", "🇹🇷"),
    "Germany": ("Alemania", "🇩🇪"), "Curacao": ("Curazao", "🇨🇼"),
    "Ivory Coast": ("Costa de Marfil", "🇨🇮"), "Ecuador": ("Ecuador", "🇪🇨"),
    "Netherlands": ("Países Bajos", "🇳🇱"), "Japan": ("Japón", "🇯🇵"),
    "Sweden": ("Suecia", "🇸🇪"), "Tunisia": ("Túnez", "🇹🇳"),
    "Belgium": ("Bélgica", "🇧🇪"), "Egypt": ("Egipto", "🇪🇬"),
    "Iran": ("Irán", "🇮🇷"), "New Zealand": ("Nueva Zelanda", "🇳🇿"),
    "Spain": ("España", "🇪🇸"), "Cape Verde": ("Cabo Verde", "🇨🇻"),
    "Uruguay": ("Uruguay", "🇺🇾"), "Saudi Arabia": ("Arabia Saudí", "🇸🇦"),
    "France": ("Francia", "🇫🇷"), "Senegal": ("Senegal", "🇸🇳"),
    "Iraq": ("Irak", "🇮🇶"), "Norway": ("Noruega", "🇳🇴"),
    "Argentina": ("Argentina", "🇦🇷"), "Algeria": ("Argelia", "🇩🇿"),
    "Austria": ("Austria", "🇦🇹"), "Jordan": ("Jordania", "🇯🇴"),
    "Portugal": ("Portugal", "🇵🇹"), "DR Congo": ("RD Congo", "🇨🇩"),
    "Uzbekistan": ("Uzbekistán", "🇺🇿"), "Colombia": ("Colombia", "🇨🇴"),
    "England": ("Inglaterra", "🏴󠁧󠁢󠁥󠁮󠁧󠁿"), "Croatia": ("Croacia", "🇭🇷"),
    "Ghana": ("Ghana", "🇬🇭"), "Panama": ("Panamá", "🇵🇦"),
}


ES2EN = {es: en for en, (es, _fl) in TEAMS.items()}

TRIVIA = {
    "Mexico": [
        ("Mexico City se hunde ~50cm al año porque está sobre un lago seco.", "Mexico City sinks ~50cm per year because it's built on a drained lake."),
        ("México regaló un axolote a la reina Isabel II en 1953.", "Mexico gifted an axolotl to Queen Elizabeth II in 1953."),
        ("El chicle fue inventado por el general mexicano Antonio López de Santa Anna.", "Chewing gum was invented by Mexican general Antonio López de Santa Anna."),
    ],
    "South Africa": [
        ("Sudáfrica tiene 11 idiomas oficiales, un récord mundial.", "South Africa has 11 official languages, a world record."),
        ("Un sudafricano inventó el CT scan.", "A South African invented the CT scan."),
        ("Table Mountain tiene más especies de plantas que todo Reino Unido.", "Table Mountain has more plant species than the entire UK."),
    ],
    "South Korea": [
        ("Corea del Sur tiene más robots per cápita que cualquier país.", "South Korea has more robots per capita than any other country."),
        ("Existe un Día del Perro Negro en Corea (3 de enero).", "There's a Black Dog Day in Korea (January 3rd)."),
        ("El servicio postal coreano entrega con drones desde 2017.", "Korea's postal service has delivered by drone since 2017."),
    ],
    "Czech Rep.": [
        ("Los checos beben más cerveza per cápita del mundo: ~140 litros/año.", "Czechs drink the most beer per capita in the world: ~140 liters/year."),
        ("La palabra 'robot' fue inventada por el escritor checo Karel Čapek en 1920.", "The word 'robot' was coined by Czech writer Karel Čapek in 1920."),
        ("Praga tiene un museo dedicado exclusivamente a Franz Kafka.", "Prague has a museum dedicated entirely to Franz Kafka."),
    ],
    "Canada": [
        ("En 1967 Canadá construyó una pista de aterrizaje para OVNIs en Alberta.", "In 1967 Canada built a UFO landing pad in Alberta."),
        ("Canadá tiene más lagos que todos los demás países juntos.", "Canada has more lakes than every other country combined."),
        ("La RCMP entrenó una vez un hipopótamo como policía montado. Fracasó.", "The RCMP once trained a hippo as a mounted police officer. It failed."),
    ],
    "Bosnia-Herz.": [
        ("Sarajevo tuvo tranvías antes que Londres.", "Sarajevo had trams before London."),
        ("Bosnia alberga la última selva virgen de Europa: Perućica.", "Bosnia is home to Europe's last primeval forest: Perućica."),
        ("Un bosnio afirma haber descubierto pirámides más antiguas que las de Egipto. Nadie le cree.", "A Bosnian claims to have discovered pyramids older than Egypt's. Nobody believes him."),
    ],
    "Qatar": [
        ("Qatar usó jinetes robot en carreras de camellos desde 2004.", "Qatar has used robot jockeys in camel races since 2004."),
        ("Qatar tiene cero impuesto sobre la renta personal.", "Qatar has zero personal income tax."),
        ("El aeropuerto de Doha tiene un oso de peluche gigante de 7 metros.", "Doha's airport features a giant 7-meter teddy bear."),
    ],
    "Switzerland": [
        ("Suiza tiene suficientes búnkeres nucleares para toda su población.", "Switzerland has enough nuclear bunkers for its entire population."),
        ("Es ilegal tener un solo conejillo de indias en Suiza (son sociales).", "It's illegal to own just one guinea pig in Switzerland (they're social)."),
        ("Suiza tiene un plan para invadir… Alemania. Por si acaso.", "Switzerland has a plan to invade… Germany. Just in case."),
    ],
    "Brazil": [
        ("Brasil tuvo un emperador que se declaró 'Protector de los Animales'.", "Brazil had an emperor who declared himself 'Protector of Animals'."),
        ("La prisión de Carandiru tenía su propia liga de fútbol oficial.", "Carandiru prison had its own official football league."),
        ("Brasil es el mayor exportador mundial de piedras preciosas.", "Brazil is the world's largest exporter of gemstones."),
    ],
    "Morocco": [
        ("Marruecos fundó la universidad más antigua del mundo aún en funcionamiento (859 d.C.).", "Morocco founded the world's oldest continuously operating university (859 AD)."),
        ("En Marruecos hay cabras que trepan a los árboles para comer argán.", "In Morocco, goats climb trees to eat argan fruit."),
        ("El rey de Marruecos tiene un trono portátil que viaja con él.", "The King of Morocco has a portable throne that travels with him."),
    ],
    "Haiti": [
        ("Haití fue el primer país negro independiente del mundo (1804).", "Haiti was the world's first independent Black republic (1804)."),
        ("Haití tiene la fortaleza más grande de América: la Citadelle Laferrière.", "Haiti has the largest fortress in the Americas: Citadelle Laferrière."),
        ("El pico más alto del Caribe está en Haití, no en República Dominicana.", "The Caribbean's highest peak is in Haiti, not the Dominican Republic."),
    ],
    "Scotland": [
        ("El animal nacional de Escocia es el unicornio.", "Scotland's national animal is the unicorn."),
        ("Escocia tiene más de 790 islas, la mayoría deshabitadas.", "Scotland has over 790 islands, most uninhabited."),
        ("Edimburgo tiene un volcán extinto en pleno centro de la ciudad.", "Edinburgh has an extinct volcano right in the city center."),
    ],
    "USA": [
        ("EE. UU. gastó 22 millones de dólares en desarrollar una pluma que escribiera en el espacio. Los rusos usaron lápiz.", "The US spent $22 million developing a pen that writes in space. The Russians used pencils."),
        ("Existe una ciudad llamada 'Hell' en Michigan. Se congela cada invierno.", "There's a town called 'Hell' in Michigan. It freezes every winter."),
        ("La bandera de EE. UU. fue diseñada por un estudiante de 17 años como proyecto escolar.", "The US flag was designed by a 17-year-old student as a school project."),
    ],
    "Paraguay": [
        ("Paraguay fue el único país latinoamericano que envió condolencias al gobierno confederado en la Guerra Civil americana.", "Paraguay was the only Latin American country to send condolences to the Confederate government during the American Civil War."),
        ("Paraguay tuvo el primer ferrocarril de Sudamérica (1854).", "Paraguay had South America's first railway (1854)."),
        ("La Armada paraguaya es la más grande del mundo… sin acceso al mar.", "Paraguay's navy is the world's largest… with no access to the sea."),
    ],
    "Australia": [
        ("Australia perdió una guerra contra 20.000 emus en 1932.", "Australia lost a war against 20,000 emus in 1932."),
        ("Australia tiene un primer ministro que fue devorado (presuntamente) por caníbales en 1803.", "Australia had a PM who was allegedly eaten by cannibals in 1803."),
        ("Hay más canguros que personas en Australia (~50M vs 26M).", "There are more kangaroos than people in Australia (~50M vs 26M)."),
    ],
    "Turkey": [
        ("Los tulipanes vienen de Turquía, no de Holanda.", "Tulips come from Turkey, not the Netherlands."),
        ("Papá Noel (San Nicolás) nació en lo que hoy es Turquía.", "Santa Claus (Saint Nicholas) was born in what is now Turkey."),
        ("Estambul tuvo un gato alcalde no oficial llamado Tombili.", "Istanbul had an unofficial cat mayor named Tombili."),
    ],
    "Germany": [
        ("Alemania tiene ~1.500 tipos de cerveza diferentes.", "Germany has ~1,500 different types of beer."),
        ("Intentaron enseñar a los perros a hablar en la Alemania nazi. No funcionó.", "They tried to teach dogs to talk in Nazi Germany. It didn't work."),
        ("Hay más clubes de fútbol que supermercados en Alemania.", "There are more football clubs than supermarkets in Germany."),
    ],
    "Curacao": [
        ("El licor azul de Curazao se hace con cáscaras de naranja de la isla.", "Blue Curaçao liqueur is made from orange peels grown on the island."),
        ("Curazao fue colonia holandesa, española, francesa e inglesa… a veces todas a la vez.", "Curaçao was a Dutch, Spanish, French, and English colony… sometimes all at once."),
        ("Curazao tiene una de las sinagogas más antiguas de América (1732).", "Curaçao has one of the oldest synagogues in the Americas (1732)."),
    ],
    "Ivory Coast": [
        ("Costa de Marfil produce el 40% del cacao mundial.", "Ivory Coast produces 40% of the world's cocoa."),
        ("Tiene una basílica más grande que la de San Pedro en el Vaticano.", "It has a basilica larger than St. Peter's in the Vatican."),
        ("El nombre 'Costa de Marfil' viene de cuando era el mayor mercado de marfil del mundo.", "The name 'Ivory Coast' comes from when it was the world's largest ivory market."),
    ],
    "Ecuador": [
        ("En la línea ecuatorial de Ecuador, el agua no hace remolino al desaguar.", "At Ecuador's equator line, water doesn't swirl when draining."),
        ("Ecuador tiene montañas donde puedes ver nieve estando en el ecuador.", "Ecuador has mountains where you can see snow while standing on the equator."),
        ("Las Islas Galápagos inspiraron la teoría de la evolución de Darwin.", "The Galápagos Islands inspired Darwin's theory of evolution."),
    ],
    "Netherlands": [
        ("Los Países Bajos están por debajo del nivel del mar en un 26% de su territorio.", "The Netherlands is below sea level in 26% of its territory."),
        ("Hay más bicicletas que personas en los Países Bajos.", "There are more bicycles than people in the Netherlands."),
        ("Holanda exportó tulipanes por primera vez como error: alguien se comió los bulbos pensando que eran cebollas.", "Holland first exported tulips by mistake: someone ate the bulbs thinking they were onions."),
    ],
    "Japan": [
        ("Japón tiene más mascotas que niños.", "Japan has more pets than children."),
        ("Hay una isla en Japón llena de conejos salvajes: Ōkunoshima.", "There's an island in Japan full of wild rabbits: Ōkunoshima."),
        ("Japón tiene un festival donde se lanzan habas a demonios imaginarios.", "Japan has a festival where you throw beans at imaginary demons."),
    ],
    "Sweden": [
        ("Suecia importa basura de otros países porque se le acabó la suya para reciclar.", "Sweden imports trash from other countries because it ran out of its own to recycle."),
        ("Hay un hotel de hielo en Suecia que se reconstruye cada invierno.", "There's an ice hotel in Sweden rebuilt every winter."),
        ("Suecia tuvo un rey que murió por comer 14 porciones de semla de una sentada.", "Sweden had a king who died from eating 14 servings of semla in one sitting."),
    ],
    "Tunisia": [
        ("Túnez fue donde se rodó Tatooine en Star Wars. Los decorados siguen ahí.", "Tunisia is where Tatooine was filmed in Star Wars. The sets are still there."),
        ("Cartago (en Túnez) tenía un puerto secreto para 220 barcos de guerra.", "Carthage (in Tunisia) had a secret harbor for 220 warships."),
        ("Túnez tiene un lago rosa natural: el Lago de Túnez.", "Tunisia has a natural pink lake: the Lake of Tunis."),
    ],
    "Belgium": [
        ("Bélgica estuvo 589 días sin gobierno (2010-2011). Nadie lo notó.", "Belgium went 589 days without a government (2010-2011). Nobody noticed."),
        ("Las patatas fritas son belgas, no francesas.", "French fries are Belgian, not French."),
        ("Bélgica tiene más cómics per cápita que cualquier otro país.", "Belgium has more comic books per capita than any other country."),
    ],
    "Egypt": [
        ("Las pirámides ya eran antiguas cuando Cleopatra vivió. Ella está más cerca de nosotros que de su construcción.", "The pyramids were already ancient when Cleopatra lived. She's closer to us than to their construction."),
        ("Los egipcios antiguos usaban maquillaje tanto hombres como mujeres.", "Ancient Egyptians wore makeup — both men and women."),
        ("Egipto tiene la presa más grande de África: la Gran Presa del Renacimiento… bueno, esa es de Etiopía. La de Asuán también es enorme.", "Egypt has one of Africa's largest dams: the Aswan High Dam."),
    ],
    "Iran": [
        ("Irán tiene más cirugías de nariz per cápita que cualquier país.", "Iran has more nose jobs per capita than any other country."),
        ("El polo fue inventado en Persia (Irán) como entrenamiento de caballería.", "Polo was invented in Persia (Iran) as cavalry training."),
        ("Irán tiene un desierto donde la temperatura alcanza los 70°C en superficie.", "Iran has a desert where surface temperatures reach 70°C."),
    ],
    "New Zealand": [
        ("Nueva Zelanda tiene más ovejas que personas (5:1).", "New Zealand has more sheep than people (5:1)."),
        ("Fue el último país grande en ser habitado por humanos (~1250 d.C.).", "It was the last major landmass settled by humans (~1250 AD)."),
        ("Nueva Zelanda tiene un pájaro que no vuela y pesa 4kg: el kiwi.", "New Zealand has a flightless bird that weighs 4kg: the kiwi."),
    ],
    "Spain": [
        ("España tiene un festival donde se lanzan tomates: La Tomatina.", "Spain has a festival where people throw tomatoes at each other: La Tomatina."),
        ("La selección española de fútbol no ganó nada importante durante 44 años… y luego ganó 3 seguidos.", "Spain's football team won nothing major for 44 years… then won 3 in a row."),
        ("España tiene más bares per cápita que cualquier país de la UE.", "Spain has more bars per capita than any EU country."),
    ],
    "Cape Verde": [
        ("Cabo Verde tiene más gente viviendo fuera del país que dentro.", "Cape Verde has more people living abroad than at home."),
        ("Es el lugar de cría más importante del mundo para tortugas marinas.", "It's the world's most important breeding site for sea turtles."),
        ("Cabo Verde no tenía población humana hasta que los portugueses llegaron en 1456.", "Cape Verde had no human population until the Portuguese arrived in 1456."),
    ],
    "Uruguay": [
        ("Uruguay tiene más vacas que personas (4:1).", "Uruguay has more cows than people (4:1)."),
        ("Uruguay fue el primer país del mundo en legalizar el cannabis a nivel nacional.", "Uruguay was the first country to legalize cannabis nationally."),
        ("El himno nacional de Uruguay dura 5 minutos. Es el más largo del mundo.", "Uruguay's national anthem lasts 5 minutes. It's the world's longest."),
    ],
    "Saudi Arabia": [
        ("Arabia Saudí importa camellos de Australia.", "Saudi Arabia imports camels from Australia."),
        ("No hay ríos permanentes en Arabia Saudí.", "There are no permanent rivers in Saudi Arabia."),
        ("Arabia Saudí tiene un edificio (Jeddah Tower) diseñado para medir 1km de alto. Aún sin terminar.", "Saudi Arabia has a building (Jeddah Tower) designed to be 1km tall. Still unfinished."),
    ],
    "France": [
        ("Francia tiene 12 husos horarios (por sus territorios de ultramar). Más que cualquier país.", "France has 12 time zones (due to overseas territories). More than any country."),
        ("Francia construyó un 'Versalles falso' para engañar a los bombarderos alemanes en la IIGM.", "France built a 'fake Versailles' to fool German bombers in WWII."),
        ("El croissant no es francés. Es austríaco.", "The croissant isn't French. It's Austrian."),
    ],
    "Senegal": [
        ("Senegal tiene un lago rosa natural: el Lago Retba.", "Senegal has a natural pink lake: Lake Retba."),
        ("El Rally Dakar originalmente terminaba en Dakar (antes de mudarse a Sudamérica).", "The Dakar Rally originally ended in Dakar (before moving to South America)."),
        ("Senegal tiene la estatua más alta de África: el Monumento al Renacimiento Africano (49m).", "Senegal has Africa's tallest statue: the African Renaissance Monument (49m)."),
    ],
    "Iraq": [
        ("Irak es donde se inventó la escritura (~3400 a.C.).", "Iraq is where writing was invented (~3400 BC)."),
        ("Las ruinas de Babilonia están a 85km de Bagdad.", "The ruins of Babylon are 85km from Baghdad."),
        ("Irak tiene los jardines colgantes de Babilonia, una de las 7 maravillas antiguas. Nunca se han encontrado.", "Iraq had the Hanging Gardens of Babylon, one of the 7 ancient wonders. They've never been found."),
    ],
    "Norway": [
        ("El rey de Noruega es un pingüino. Literalmente. Se llama Nils Olav.", "The King of Norway is a penguin. Literally. His name is Nils Olav."),
        ("Noruega regaló un árbol de Navidad a Londres cada año desde 1947. Como agradecimiento por la WWII.", "Norway has gifted a Christmas tree to London every year since 1947. As thanks for WWII."),
        ("Noruega tiene un pueblo fantasma pirámide en el Ártico: Pyramiden.", "Norway has a ghost town pyramid in the Arctic: Pyramiden."),
    ],
    "Argentina": [
        ("Argentina tuvo 5 presidentes en 10 días en 2001.", "Argentina had 5 presidents in 10 days in 2001."),
        ("El tango nació en los burdeles de Buenos Aires.", "Tango was born in the brothels of Buenos Aires."),
        ("Argentina tiene la avenida más ancha del mundo: la 9 de Julio (14 carriles).", "Argentina has the world's widest avenue: 9 de Julio (14 lanes)."),
    ],
    "Algeria": [
        ("Argelia es el país más grande de África por superficie.", "Algeria is the largest country in Africa by area."),
        ("Argelia tiene más de 1.000 km de costa mediterránea.", "Algeria has over 1,000 km of Mediterranean coastline."),
        ("El desierto del Sahara cubre el 80% de Argelia.", "The Sahara Desert covers 80% of Algeria."),
    ],
    "Austria": [
        ("Austria tiene un pueblo llamado 'Fucking'. Tuvo que cambiar el nombre por los turistas.", "Austria had a village called 'Fucking'. It had to change the name because of tourists."),
        ("El 80% de la banda sonora de 'The Sound of Music' se rodó en Austria.", "80% of 'The Sound of Music' was filmed in Austria."),
        ("Austria inventó la bola de nieve (Schneekugel) en 1900.", "Austria invented the snow globe (Schneekugel) in 1900."),
    ],
    "Jordan": [
        ("Jordania tiene el punto más bajo de la Tierra: el Mar Muerto (-430m).", "Jordan has the lowest point on Earth: the Dead Sea (-430m)."),
        ("Petra (en Jordania) fue tallada en roca rosa hace 2.000 años.", "Petra (in Jordan) was carved into pink rock 2,000 years ago."),
        ("Jordania tiene un castillo en medio del desierto que parece una nave espacial: Qasr Kharana.", "Jordan has a desert castle that looks like a spaceship: Qasr Kharana."),
    ],
    "Portugal": [
        ("Portugal tiene la alianza diplomática más antigua del mundo con Inglaterra (1373).", "Portugal has the world's oldest diplomatic alliance with England (1373)."),
        ("Portugal es el país más antiguo de Europa con las mismas fronteras desde 1139.", "Portugal is Europe's oldest country with the same borders since 1139."),
        ("Lisboa es más antigua que Roma.", "Lisbon is older than Rome."),
    ],
    "DR Congo": [
        ("La RD Congo tiene el 50% de las reservas mundiales de cobalto.", "DR Congo has 50% of the world's cobalt reserves."),
        ("El río Congo es el más profundo del mundo (220m).", "The Congo River is the world's deepest (220m)."),
        ("La RD Congo tiene un volcán que produce lava azul: Nyiragongo.", "DR Congo has a volcano that produces blue lava: Nyiragongo."),
    ],
    "Uzbekistan": [
        ("Uzbekistán tiene una de las ciudades más antiguas de la Ruta de la Seda: Samarcanda.", "Uzbekistan has one of the oldest Silk Road cities: Samarkand."),
        ("El Mar de Aral (en Uzbekistán) era el 4º lago más grande del mundo. Ahora es un desierto.", "The Aral Sea (in Uzbekistan) was the world's 4th largest lake. Now it's a desert."),
        ("Uzbekistán tiene el metro más bonito de Asia Central. Cada estación es una obra de arte.", "Uzbekistan has Central Asia's most beautiful metro. Each station is a work of art."),
    ],
    "Colombia": [
        ("Colombia tiene más especies de aves que cualquier otro país (~1.900).", "Colombia has more bird species than any other country (~1,900)."),
        ("El río Caño Cristales en Colombia tiene 5 colores y le llaman 'el río que se escapó del paraíso'.", "Colombia's Caño Cristales river has 5 colors and is called 'the river that escaped paradise'."),
        ("Colombia es el único país de Sudamérica con costa en el Pacífico Y el Caribe.", "Colombia is the only South American country with both Pacific and Caribbean coasts."),
    ],
    "England": [
        ("Inglaterra tuvo una reina que reinó solo 9 días: Lady Jane Grey.", "England had a queen who reigned just 9 days: Lady Jane Grey."),
        ("El Big Ben no es el nombre de la torre. Es el nombre de la campana.", "Big Ben isn't the tower's name. It's the bell's name."),
        ("Inglaterra inventó el fútbol moderno… y luego no ganó un Mundial durante 56 años.", "England invented modern football… then didn't win a World Cup for 56 years."),
    ],
    "Croatia": [
        ("La corbata fue inventada por mercenarios croatas en el siglo XVII.", "The necktie was invented by Croatian mercenaries in the 17th century."),
        ("Croacia tiene el museo de relaciones rotas del mundo (Zagreb).", "Croatia has the world's Museum of Broken Relationships (Zagreb)."),
        ("Dálmata (el perro) viene de Dalmacia, Croacia.", "The Dalmatian dog comes from Dalmatia, Croatia."),
    ],
    "Ghana": [
        ("Ghana tiene el ataúd más customizado del mundo: los hacen con forma de pez, avión, Coca-Cola…", "Ghana has the world's most custom coffins: shaped like fish, planes, Coca-Cola bottles…"),
        ("Ghana fue llamada 'Costa de Oro' por los colonizadores portugueses.", "Ghana was called the 'Gold Coast' by Portuguese colonizers."),
        ("El lago Volta en Ghana es el mayor lago artificial del mundo por superficie.", "Lake Volta in Ghana is the world's largest artificial lake by surface area."),
    ],
    "Panama": [
        ("Panamá es el único lugar del mundo donde puedes ver el amanecer en el Pacífico y el atardecer en el Atlántico.", "Panama is the only place where you can see sunrise on the Pacific and sunset on the Atlantic."),
        ("El Canal de Panamá mueve el 6% del comercio mundial.", "The Panama Canal moves 6% of world trade."),
        ("Panamá tiene más de 1.500 islas. La mayoría sin nombre.", "Panama has over 1,500 islands. Most unnamed."),
    ],
}


def team_es(name):
    if name is None:
        return None
    name = str(name).strip()
    return TEAMS.get(name, (name, "🏳️"))[0]


def team_flag(name):
    if name is None:
        return "🏳️"
    return TEAMS.get(str(name).strip(), (name, "🏳️"))[1]


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------

def parse_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    raw = wb["Raw data"]

    # Localiza columnas de participantes con nombre en NAME_ROW
    participants = []  # [{name, home_col, away_col, idx}]
    for slot in range(N_PARTICIPANTS_MAX):
        home_col = RAW_FIRST_HOME_COL + 2 * slot
        name = raw.cell(NAME_ROW, home_col).value
        if name is not None and str(name).strip():
            participants.append({
                "name": str(name).strip(),
                "home_col": home_col,
                "away_col": home_col + 1,
                "idx": len(participants),
            })
    names = [p["name"] for p in participants]
    n = len(participants)

    # Partidos de fase de grupos
    matches = []
    for r in GROUP_MATCH_ROWS:
        concept = raw.cell(r, 2).value
        if not concept or "vs" not in str(concept):
            continue
        concept = str(concept).strip()
        code, _, teams = concept.partition(":")
        code = code.strip()
        home_en, _, away_en = teams.strip().partition(" vs ")
        home_en, away_en = home_en.strip(), away_en.strip()
        group = code.split("-")[0].replace("G", "")  # "GA-M1" -> "A"
        picks = []
        for p in participants:
            hg = raw.cell(r, p["home_col"]).value
            ag = raw.cell(r, p["away_col"]).value
            picks.append((_num(hg), _num(ag)))
        matches.append({
            "code": code, "group": group,
            "home_en": home_en, "away_en": away_en,
            "home": team_es(home_en), "away": team_es(away_en),
            "home_flag": team_flag(home_en), "away_flag": team_flag(away_en),
            "picks": picks,
            "date": MATCH_DATES.get(code, ""),
        })

    # Clasificados por grupo (1/2/3)
    qualifiers = defaultdict(dict)  # group -> {pos -> [team per participant]}
    for r in QUALIFIER_ROWS:
        label = raw.cell(r, 2).value
        if not label:
            continue
        label = str(label).strip().lower()
        pos = None
        if label.startswith("1st"):
            pos = 1
        elif label.startswith("2nd"):
            pos = 2
        elif label.startswith("3rd"):
            pos = 3
        if pos is None:
            continue
        group = label.split("group")[-1].strip().upper()
        picks = [raw.cell(r, p["home_col"]).value for p in participants]
        qualifiers[group][pos] = picks

    # 8 mejores terceros
    thirds = []  # list per row -> [team per participant]
    for r in THIRDS_ROWS:
        label = raw.cell(r, 2).value
        if not label:
            continue
        picks = [raw.cell(r, p["home_col"]).value for p in participants]
        thirds.append(picks)

    # ¿Hay resultados reales?
    results = parse_results(wb)

    return {
        "participants": participants,
        "names": names,
        "n": n,
        "matches": matches,
        "qualifiers": qualifiers,
        "thirds": thirds,
        "results": results,
    }


def parse_results(wb):
    """Lee resultados reales de fase de grupos si existen. Devuelve {code:(h,a)}."""
    rr = wb["Real results"]
    out = {}
    for r in range(1, rr.max_row + 1):
        concept = rr.cell(r, 2).value
        if not concept or "vs" not in str(concept):
            continue
        code = str(concept).split(":")[0].strip()
        h = _num(rr.cell(r, 3).value)
        a = _num(rr.cell(r, 4).value)
        if h is not None and a is not None:
            out[code] = (h, a)
    return out


def _num(v):
    if v is None or v == "":
        return None
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------
# Utilidades de analítica
# --------------------------------------------------------------------------

def outcome(hg, ag):
    if hg is None or ag is None:
        return None
    if hg > ag:
        return "1"
    if hg < ag:
        return "2"
    return "X"


def minmax_scale(values):
    lo, hi = min(values), max(values)
    if hi - lo < 1e-9:
        return [50.0 for _ in values]
    return [100.0 * (v - lo) / (hi - lo) for v in values]


def zscore(values):
    mu = statistics.mean(values)
    sd = statistics.pstdev(values) or 1.0
    return [(v - mu) / sd for v in values]


# --------------------------------------------------------------------------
# Motor de analíticas
# --------------------------------------------------------------------------

def compute(data):
    names = data["names"]
    n = data["n"]
    matches = data["matches"]
    nm = len(matches)

    # ---- Índice de Rebeldía (signo minoría + distancia marcador) ----
    raw_sign = [0.0] * n
    raw_score = [0.0] * n
    match_meta = []  # por partido: agreement, mediana, etc.

    for m in matches:
        picks = m["picks"]
        valid = [(i, h, a) for i, (h, a) in enumerate(picks) if h is not None and a is not None]
        if not valid:
            match_meta.append(None)
            continue
        outcomes = [outcome(h, a) for _, h, a in valid]
        oc = Counter(outcomes)
        total = len(valid)
        homes = [h for _, h, a in valid]
        aways = [a for _, h, a in valid]
        med_h = statistics.median(homes)
        med_a = statistics.median(aways)
        # distancia de marcador por persona
        sdist = {}
        for i, h, a in valid:
            sdist[i] = abs(h - med_h) + abs(a - med_a)
        max_sdist = max(sdist.values()) or 1.0
        for i, h, a in valid:
            o = outcome(h, a)
            share = oc[o] / total
            raw_sign[i] += (1.0 - share)
            raw_score[i] += sdist[i] / max_sdist
        # scoreline modal
        scl = Counter((h, a) for _, h, a in valid)
        modal_scl, modal_cnt = scl.most_common(1)[0]
        top_outcome, top_cnt = oc.most_common(1)[0]
        match_meta.append({
            "code": m["code"], "group": m["group"],
            "home": m["home"], "away": m["away"],
            "home_en": m["home_en"], "away_en": m["away_en"],
            "home_flag": m["home_flag"], "away_flag": m["away_flag"],
            "outcome_dist": dict(oc),
            "outcome_agreement": top_cnt / total,
            "top_outcome": top_outcome,
            "modal_scoreline": f"{modal_scl[0]}-{modal_scl[1]}",
            "modal_scoreline_share": modal_cnt / total,
            "median": f"{int(med_h)}-{int(med_a)}",
            "n": total,
        })

    raw_sign = [s / nm for s in raw_sign]
    raw_score = [s / nm for s in raw_score]
    z = [0.5 * a + 0.5 * b for a, b in zip(zscore(raw_sign), zscore(raw_score))]
    rebel_index = [round(v, 1) for v in minmax_scale(z)]

    rebeldia = sorted(
        [{"name": names[i], "index": rebel_index[i],
          "sign": round(raw_sign[i], 3), "score": round(raw_score[i], 3)}
         for i in range(n)],
        key=lambda x: -x["index"],
    )

    # ---- Similitud: % marcadores idénticos + distancia media de goles ----
    exact = [[0] * n for _ in range(n)]
    gdist = [[0.0] * n for _ in range(n)]
    comparable = 0
    for m in matches:
        picks = m["picks"]
        if all(h is not None for h, a in picks):
            comparable += 1
        for i in range(n):
            hi, ai = picks[i]
            if hi is None:
                continue
            for j in range(i + 1, n):
                hj, aj = picks[j]
                if hj is None:
                    continue
                if hi == hj and ai == aj:
                    exact[i][j] += 1
                    exact[j][i] += 1
                d = abs(hi - hj) + abs(ai - aj)
                gdist[i][j] += d
                gdist[j][i] += d

    sim_pct = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(n):
            if i != j and comparable:
                sim_pct[i][j] = round(100.0 * exact[i][j] / comparable, 1)

    # almas gemelas (máx % idéntico) y polos opuestos (máx distancia)
    best_pair = (None, None, -1)
    worst_pair = (None, None, -1)
    for i in range(n):
        for j in range(i + 1, n):
            if sim_pct[i][j] > best_pair[2]:
                best_pair = (i, j, sim_pct[i][j])
            avg_d = gdist[i][j] / (comparable or 1)
            if avg_d > worst_pair[2]:
                worst_pair = (i, j, round(avg_d, 2))

    # gemelo de cada uno
    twins = []
    for i in range(n):
        best_j, best_v = None, -1
        for j in range(n):
            if i != j and sim_pct[i][j] > best_v:
                best_v, best_j = sim_pct[i][j], j
        twins.append({"name": names[i], "twin": names[best_j], "pct": best_v})

    # ---- Estilo de apuesta ----
    style = []
    all_scorelines = Counter()
    for i in range(n):
        goals, draws, blowouts, biggest = [], 0, 0, (None, -1)
        for m in matches:
            h, a = m["picks"][i]
            if h is None:
                continue
            goals.append(h + a)
            if h == a:
                draws += 1
            if abs(h - a) >= 3:
                blowouts += 1
            if h + a > biggest[1]:
                biggest = ((m, h, a), h + a)
            all_scorelines[(h, a)] += 1
        cnt = len(goals) or 1
        big = None
        if biggest[0]:
            mm, h, a = biggest[0]
            big = {"home": mm["home"], "away": mm["away"],
                   "home_en": mm["home_en"], "away_en": mm["away_en"],
                   "score": f"{h}-{a}",
                   "flags": f"{mm['home_flag']}{mm['away_flag']}", "goals": h + a}
        style.append({
            "name": names[i],
            "avg_goals": round(sum(goals) / cnt, 2),
            "pct_draws": round(100 * draws / cnt, 1),
            "pct_blowouts": round(100 * blowouts / cnt, 1),
            "biggest": big,
        })

    common_scorelines = [
        {"score": f"{h}-{a}", "count": c,
         "pct": round(100 * c / (nm * n), 1)}
        for (h, a), c in all_scorelines.most_common(8)
    ]

    # ---- Lobo solitario: marcadores únicos (solo 1 persona) ----
    unique_count = [0] * n
    unique_examples = []  # picks únicos llamativos
    for m in matches:
        scl = Counter()
        for h, a in m["picks"]:
            if h is not None:
                scl[(h, a)] += 1
        for i, (h, a) in enumerate(m["picks"]):
            if h is not None and scl[(h, a)] == 1:
                unique_count[i] += 1
                diff = abs(h - a) + (h + a)
                unique_examples.append({
                    "name": names[i], "home": m["home"], "away": m["away"],
                    "home_en": m["home_en"], "away_en": m["away_en"],
                    "flags": f"{m['home_flag']}{m['away_flag']}",
                    "score": f"{h}-{a}", "spice": diff,
                })
    lobo = sorted(
        [{"name": names[i], "count": unique_count[i]} for i in range(n)],
        key=lambda x: -x["count"],
    )
    unique_examples.sort(key=lambda x: -x["spice"])

    # ---- Favoritos del torneo ----
    group_consensus = []
    respect = Counter()        # prestigio ponderado: 1º=3, 2º=2, 3º=1
    pos_weight = {1: 3, 2: 2, 3: 1}
    for group in sorted(data["qualifiers"].keys()):
        pos = data["qualifiers"][group]
        firsts = [team_es(t) for t in pos.get(1, []) if t]
        c = Counter(firsts)
        for p, slot_picks in pos.items():
            w = pos_weight.get(p, 1)
            for t in slot_picks:
                if t:
                    respect[team_es(t)] += w
        if c:
            top, cnt = c.most_common(1)[0]
            group_consensus.append({
                "group": group, "favorite": top,
                "flag": _flag_es(top),
                "agreement": round(100 * cnt / sum(c.values()), 1),
                "dist": dict(c.most_common()),
            })
    most_respected = [{"team": t, "flag": _flag_es(t), "count": c}
                      for t, c in respect.most_common(12)]
    least_respected = [{"team": t, "flag": _flag_es(t), "count": c}
                       for t, c in sorted(respect.items(), key=lambda x: x[1])[:8]]
    polarizing = sorted(group_consensus, key=lambda x: x["agreement"])[:5]

    # terceros más elegidos
    thirds_counter = Counter()
    for row in data["thirds"]:
        for t in row:
            if t:
                thirds_counter[team_es(t)] += 1
    top_thirds = [{"team": t, "flag": _flag_es(t), "count": c}
                  for t, c in thirds_counter.most_common(10)]

    # ---- Partidos divisivos / unánimes ----
    valid_meta = [m for m in match_meta if m]
    divisive = sorted(valid_meta, key=lambda m: (m["outcome_agreement"], m["modal_scoreline_share"]))[:6]
    unanimous = sorted(valid_meta, key=lambda m: (-m["modal_scoreline_share"], -m["outcome_agreement"]))[:6]

    # ---- Fichas por persona ----
    rank_of = {r["name"]: i + 1 for i, r in enumerate(rebeldia)}
    idx_of = {r["name"]: r["index"] for r in rebeldia}
    style_of = {s["name"]: s for s in style}
    twin_of = {t["name"]: t for t in twins}
    lobo_of = {l["name"]: l["count"] for l in lobo}
    cards = []
    for i, name in enumerate(names):
        st = style_of[name]
        cards.append({
            "name": name,
            "rebel_rank": rank_of[name],
            "rebel_index": idx_of[name],
            "avg_goals": st["avg_goals"],
            "pct_draws": st["pct_draws"],
            "biggest": st["biggest"],
            "twin": twin_of[name]["twin"],
            "twin_pct": twin_of[name]["pct"],
            "lobo": lobo_of[name],
            "label": style_label(st, idx_of[name]),
        })

    # ---- Premios ----
    goleador = max(style, key=lambda s: s["avg_goals"])
    cerrojo = min(style, key=lambda s: s["avg_goals"])
    empate = max(style, key=lambda s: s["pct_draws"])
    awards = {
        "rebelde": rebeldia[0],
        "borrego": rebeldia[-1],
        "gemelas": {"a": names[best_pair[0]], "b": names[best_pair[1]], "pct": best_pair[2]},
        "opuestos": {"a": names[worst_pair[0]], "b": names[worst_pair[1]], "dist": worst_pair[2]},
        "goleador": {"name": goleador["name"], "avg": goleador["avg_goals"]},
        "cerrojo": {"name": cerrojo["name"], "avg": cerrojo["avg_goals"]},
        "empate": {"name": empate["name"], "pct": empate["pct_draws"]},
        "lobo": lobo[0],
    }

    # ---- Stats del hero ----
    total_goals = sum(
        h + a for m in matches for (h, a) in m["picks"] if h is not None
    )
    hero = {
        "participants": n,
        "matches": nm,
        "groups": len(data["qualifiers"]),
        "total_goals": total_goals,
        "avg_goals_match": round(total_goals / (nm * n), 2) if nm and n else 0,
        "top_scoreline": common_scorelines[0] if common_scorelines else None,
        "has_results": bool(data["results"]),
    }

    # ---- Aciertos en directo (solo si hay resultados) ----
    live = None
    if data["results"]:
        live = compute_live(data, matches)

    today = compute_today(data, matches)
    recent_results = compute_recent_results(data, matches)

    return {
        "es2en": ES2EN,
        "hero": hero,
        "rebeldia": rebeldia,
        "matrix": {"names": names, "sim": sim_pct},
        "twins": twins,
        "best_pair": {"a": names[best_pair[0]], "b": names[best_pair[1]], "pct": best_pair[2]},
        "worst_pair": {"a": names[worst_pair[0]], "b": names[worst_pair[1]], "dist": worst_pair[2]},
        "style": style,
        "common_scorelines": common_scorelines,
        "lobo": lobo,
        "unique_examples": unique_examples[:12],
        "group_consensus": group_consensus,
        "most_respected": most_respected,
        "least_respected": least_respected,
        "polarizing": polarizing,
        "top_thirds": top_thirds,
        "divisive": divisive,
        "unanimous": unanimous,
        "cards": cards,
        "awards": awards,
        "live": live,
        "today": today,
        "recent_results": recent_results,
    }


def _flag_es(es_name):
    for en, (es, fl) in TEAMS.items():
        if es == es_name:
            return fl
    return "🏳️"


def style_label(st, rebel_index):
    """Devuelve una clave; el texto/emoji se traduce en el cliente (JS)."""
    if st["avg_goals"] >= 3.0:
        return "goleador"
    if st["avg_goals"] <= 2.0:
        return "cerrojo"
    if st["pct_draws"] >= 35:
        return "empate"
    if rebel_index >= 70:
        return "rebelde"
    if rebel_index <= 25:
        return "borrego"
    return "equilibrado"


def compute_live(data, matches):
    """Ranking simplificado de aciertos cuando hay resultados (signo=2, pleno=4)."""
    results = data["results"]
    names = data["names"]
    n = data["n"]
    pts = [0] * n
    exact = [0] * n
    sign = [0] * n
    name_index = {name: i for i, name in enumerate(names)}

    def ranked_rows():
        rows = [
            {
                "name": names[i],
                "pts": pts[i],
                "exact": exact[i],
                "sign": sign[i],
                "_order": i,
            }
            for i in range(n)
        ]
        rows.sort(key=lambda x: (-x["pts"], x["_order"]))
        for rank, row in enumerate(rows, 1):
            row["rank"] = rank
            del row["_order"]
        return rows

    played_matches = sorted(
        [m for m in matches if m["code"] in results],
        key=lambda m: (m["date"], m["code"]),
    )
    progression = []
    previous_ranks = {}
    for idx, m in enumerate(played_matches, 1):
        before_pts = pts[:]
        before_exact = exact[:]
        before_sign = sign[:]
        rh, ra = results[m["code"]]
        ro = outcome(rh, ra)
        for i, (h, a) in enumerate(m["picks"]):
            if h is None:
                continue
            if h == rh and a == ra:
                pts[i] += 4
                exact[i] += 1
            elif outcome(h, a) == ro:
                pts[i] += 2
                sign[i] += 1
        rows = ranked_rows()
        for row in rows:
            i = name_index[row["name"]]
            old_rank = previous_ranks.get(row["name"], row["rank"])
            row["delta"] = old_rank - row["rank"]
            row["round_pts"] = pts[i] - before_pts[i]
            row["round_exact"] = exact[i] - before_exact[i]
            row["round_sign"] = sign[i] - before_sign[i]
        previous_ranks = {row["name"]: row["rank"] for row in rows}
        progression.append({
            "idx": idx,
            "code": m["code"],
            "group": m["group"],
            "date": m["date"],
            "home": m["home"],
            "away": m["away"],
            "home_en": m["home_en"],
            "away_en": m["away_en"],
            "home_flag": m["home_flag"],
            "away_flag": m["away_flag"],
            "result": {"home": rh, "away": ra, "outcome": ro},
            "table": rows,
        })
    table = progression[-1]["table"] if progression else ranked_rows()
    return {"played": len(played_matches), "table": table, "progression": progression}


def compute_recent_results(data, matches, limit=6):
    """Últimos partidos con resultado real y quién acertó/falló."""
    results = data["results"]
    names = data["names"]
    played = []
    for m in matches:
        if m["code"] not in results:
            continue
        rh, ra = results[m["code"]]
        ro = outcome(rh, ra)
        exact = []
        sign = []
        miss = []
        for name, (h, a) in zip(names, m["picks"]):
            if h is None or a is None:
                miss.append({"name": name, "pick": "–"})
            elif h == rh and a == ra:
                exact.append({"name": name, "pick": f"{h}-{a}"})
            elif outcome(h, a) == ro:
                sign.append({"name": name, "pick": f"{h}-{a}"})
            else:
                miss.append({"name": name, "pick": f"{h}-{a}"})
        played.append({
            "code": m["code"], "group": m["group"], "date": m["date"],
            "home_en": m["home_en"], "away_en": m["away_en"],
            "home": m["home"], "away": m["away"],
            "home_flag": m["home_flag"], "away_flag": m["away_flag"],
            "result": {"home": rh, "away": ra, "outcome": ro},
            "exact": exact,
            "sign": sign,
            "miss": miss,
        })
    played.sort(key=lambda m: (m["date"], m["code"]), reverse=True)
    return {"matches": played[:limit], "total": len(played)}


def compute_today(data, matches):
    names = data["names"]
    all_matches = []
    trivia_used = defaultdict(int)
    for m in matches:
        picks = []
        for i, (h, a) in enumerate(m["picks"]):
            picks.append({"name": names[i], "home": h, "away": a})
        o1 = sum(1 for h, a in m["picks"] if h is not None and h > a)
        ox = sum(1 for h, a in m["picks"] if h is not None and h == a)
        o2 = sum(1 for h, a in m["picks"] if h is not None and h < a)
        cnt = Counter(f"{h}-{a}" for h, a in m["picks"] if h is not None)
        modal = cnt.most_common(1)
        modal_sl = modal[0][0] if modal else "–"
        modal_pct = round(modal[0][1] / len(m["picks"]) * 100, 1) if modal else 0
        unique_picks = [(n, f"{h}-{a}") for n, (h, a) in zip(names, m["picks"])
                        if h is not None and cnt.get(f"{h}-{a}", 0) == 1]
        most_unique = None
        if unique_picks:
            best = max(unique_picks, key=lambda x: sum(int(g) for g in x[1].split("-")) + abs(int(x[1].split("-")[0]) - int(x[1].split("-")[1])))
            most_unique = {"name": best[0], "score": best[1]}
        t_idx_home = trivia_used[m["home_en"]]
        t_idx_away = trivia_used[m["away_en"]]
        trivia_used[m["home_en"]] += 1
        trivia_used[m["away_en"]] += 1
        home_trivia = TRIVIA.get(m["home_en"], [("", "")] * 3)[t_idx_home % 3]
        away_trivia = TRIVIA.get(m["away_en"], [("", "")] * 3)[t_idx_away % 3]
        all_matches.append({
            "code": m["code"], "group": m["group"], "date": m["date"],
            "home_en": m["home_en"], "away_en": m["away_en"],
            "home": m["home"], "away": m["away"],
            "home_flag": m["home_flag"], "away_flag": m["away_flag"],
            "picks": picks,
            "outcome_dist": {"1": o1, "X": ox, "2": o2},
            "modal_scoreline": modal_sl,
            "modal_scoreline_share": modal_pct / 100,
            "most_unique_pick": most_unique,
            "home_trivia": {"es": home_trivia[0], "en": home_trivia[1]},
            "away_trivia": {"es": away_trivia[0], "en": away_trivia[1]},
        })
    return {"matches": all_matches}


# --------------------------------------------------------------------------
# main (parte de verificación; el render se añade después)
# --------------------------------------------------------------------------

CSS = r"""
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#022f36;--bg2:#001a1f;--teal:#003C46;--surface:#063a44;--surface2:#0a4a56;
  --mint:#7afcd0;--mint2:#49d6bb;--mint-soft:#defff0;--line:rgba(122,252,208,.15);
  --text:#e2f6ef;--muted:#86b1aa;--gold:#ffd27a;--red:#ff7a7a;--maxw:1060px;
}
html{scroll-behavior:smooth}
body{background:radial-gradient(1200px 800px at 75% -10%,#0a4a56 0%,var(--bg) 45%,var(--bg2) 100%);
  color:var(--text);font-family:'Inter',system-ui,-apple-system,Segoe UI,Roboto,sans-serif;
  line-height:1.5;-webkit-font-smoothing:antialiased;min-height:100vh;overflow-x:hidden}
h1,h2,h3,.disp{font-family:'Space Grotesk','Inter',system-ui,sans-serif;font-weight:700;letter-spacing:-.01em;line-height:1.05}
a{color:inherit}
.mint{color:var(--mint)}.muted{color:var(--muted)}
/* NAV */
nav.rail{position:fixed;top:0;left:0;height:100vh;width:188px;padding:26px 18px;display:flex;flex-direction:column;gap:4px;
  border-right:1px solid var(--line);background:linear-gradient(180deg,rgba(0,26,31,.6),transparent);backdrop-filter:blur(6px);z-index:40;overflow-y:auto}
nav.rail .brand{color:var(--mint);height:26px;margin-bottom:18px;display:block}
nav.rail .brand svg{height:26px;width:auto}
.langtoggle{display:flex;gap:4px;margin-bottom:16px;background:rgba(0,0,0,.2);border:1px solid var(--line);border-radius:10px;padding:3px}
.langtoggle button{flex:1;border:0;background:transparent;color:var(--muted);font:600 .78rem 'Space Grotesk',sans-serif;
  padding:6px 0;border-radius:7px;cursor:pointer;letter-spacing:.05em;transition:.18s}
.langtoggle button.on{background:var(--mint);color:#012}
nav.rail a{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:10px;color:var(--muted);
  text-decoration:none;font-size:.82rem;font-weight:600;transition:.2s}
nav.rail a .dot{width:7px;height:7px;border-radius:50%;background:currentColor;opacity:.5;transition:.2s}
nav.rail a:hover{color:var(--text);background:rgba(122,252,208,.05)}
nav.rail a.active{color:var(--mint);background:rgba(122,252,208,.1)}
nav.rail a.active .dot{opacity:1;box-shadow:0 0 10px var(--mint)}
.nav-group-label{color:rgba(226,246,239,.54);font-size:.64rem;font-weight:700;text-transform:uppercase;letter-spacing:.11em;margin:12px 10px 2px}
.wrap{max-width:var(--maxw);margin:0 auto;padding:0 28px 0 calc(188px + 40px)}
/* HERO */
header.hero{min-height:100vh;display:flex;flex-direction:column;justify-content:center;position:relative;padding-top:40px}
.hero .eyebrow{color:var(--mint);font-weight:700;letter-spacing:.22em;text-transform:uppercase;font-size:.78rem;margin-bottom:18px}
.hero h1{font-size:clamp(2.6rem,8vw,5.6rem);background:linear-gradient(120deg,#fff 10%,var(--mint) 90%);
  -webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;margin-bottom:10px}
.hero .lead{font-size:clamp(1rem,2vw,1.3rem);color:var(--muted);max-width:620px;margin-bottom:42px}
.chips{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;max-width:920px}
.chip{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:18px 22px;min-width:0}
.chip .big{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.chip .big{font-family:'Space Grotesk';font-size:2rem;font-weight:700;color:var(--mint)}
.chip .lab{font-size:.78rem;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-top:4px}
.scrollcue{position:absolute;bottom:26px;left:0;color:var(--muted);font-size:.8rem;display:flex;align-items:center;gap:8px;animation:bob 1.8s ease-in-out infinite}
@keyframes bob{0%,100%{transform:translateY(0)}50%{transform:translateY(6px)}}
/* SECTIONS */
section.sec{padding:74px 0;border-top:1px solid var(--line)}
.sec-head{margin-bottom:34px}
.kicker{font-family:'Space Grotesk';color:var(--mint);font-weight:700;font-size:.9rem;letter-spacing:.1em;margin-bottom:8px}
.sec h2{font-size:clamp(1.8rem,4vw,2.9rem)}
.sec .sub{color:var(--muted);max-width:640px;margin-top:10px;font-size:1.02rem}
.grid{display:grid;gap:16px}
.g2{grid-template-columns:repeat(2,1fr)}.g3{grid-template-columns:repeat(3,1fr)}.g4{grid-template-columns:repeat(4,1fr)}
.card{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:22px}
.card.glow{box-shadow:0 0 0 1px rgba(122,252,208,.2),0 20px 60px rgba(122,252,208,.08)}
.card h3{font-size:1.05rem;margin-bottom:4px}
.card .k{font-size:.74rem;color:var(--muted);text-transform:uppercase;letter-spacing:.1em}
.big-num{font-family:'Space Grotesk';font-size:2.4rem;font-weight:700;color:var(--mint);line-height:1}
/* BARS */
.bar-row{display:grid;grid-template-columns:30px minmax(84px,170px) 1fr 54px;align-items:center;gap:14px;padding:7px 0}
.bar-rank{font-family:'Space Grotesk';color:var(--muted);font-size:.95rem;text-align:right;font-weight:700}
.bar-name{font-weight:600;font-size:.95rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar-co{display:flex;flex-direction:column;gap:3px;min-width:0}
.bar-track{height:13px;background:rgba(255,255,255,.06);border-radius:8px;overflow:hidden}
.bar-fill{height:100%;width:0;border-radius:8px;background:linear-gradient(90deg,var(--mint2),var(--mint));transition:width 1s cubic-bezier(.2,.8,.2,1)}
.bar-fill.cool{background:linear-gradient(90deg,#2b6f7f,#4a9aa8)}
.bar-fill.gold{background:linear-gradient(90deg,#caa24a,var(--gold))}
.bar-val{font-family:'Space Grotesk';font-weight:700;text-align:right;font-size:.95rem}
/* ranking evolution */
.rank-proto-card{background:rgba(0,26,31,.34);border:1px solid var(--line);border-radius:18px;padding:18px;margin-top:18px;overflow:hidden}
.rank-proto-top{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;margin-bottom:14px}
.rank-proto-top h3{font-size:1.08rem;margin-bottom:4px}
.rank-proto-meta{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end}
.rank-state-pill{background:rgba(255,255,255,.06);border:1px solid var(--line);border-radius:999px;padding:6px 9px;
  color:var(--muted);font-size:.78rem;white-space:nowrap}
.rank-state-pill b{color:var(--text);font-family:'Space Grotesk'}
.bump-scroll,.rank-matrix-scroll{overflow-x:auto;padding:6px 2px 12px}
.bump-svg{display:block;min-width:880px}
.bump-grid{stroke:rgba(226,246,239,.13);stroke-dasharray:2 5}
.bump-axis-label{fill:var(--muted);font:700 11px 'Space Grotesk',sans-serif}
.bump-rank-label{fill:rgba(226,246,239,.6);font:700 11px 'Space Grotesk',sans-serif}
.bump-name{font:700 12px 'Inter',sans-serif}
.bump-score{fill:var(--muted);font:600 11px 'Space Grotesk',sans-serif}
.bump-line{fill:none;stroke-linecap:round;stroke-linejoin:round}
.bump-line-muted{opacity:.22}
.bump-point{stroke:#001a1f;stroke-width:2}
.rank-table-compact{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:8px;margin-top:14px}
.rank-table-compact .rank-row{display:grid;grid-template-columns:28px minmax(0,1fr) auto;gap:8px;align-items:center;background:rgba(255,255,255,.045);
  border:1px solid rgba(122,252,208,.1);border-radius:10px;padding:8px 10px;font-size:.83rem}
.rank-row .rr-name{font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rank-row .rr-points{font-family:'Space Grotesk';font-weight:700;color:var(--gold)}
.rank-delta{font-family:'Space Grotesk';font-size:.72rem;color:var(--muted);white-space:nowrap}
.rank-delta.up{color:var(--mint)}.rank-delta.down{color:var(--red)}
.race-layout{display:grid;grid-template-columns:minmax(220px,280px) 1fr;gap:16px;align-items:start}
.race-control{background:linear-gradient(160deg,rgba(122,252,208,.08),rgba(255,210,122,.055));border:1px solid var(--line);border-radius:14px;padding:14px}
.race-control input{width:100%;accent-color:var(--mint);margin:12px 0}
.race-actions{display:grid;grid-template-columns:36px 36px 1fr 36px;gap:8px;align-items:center;margin:12px 0}
.race-actions button{height:36px;border:1px solid var(--line);border-radius:50%;background:#001a1f;color:var(--mint);
  font:800 .98rem 'Space Grotesk';cursor:pointer}
.race-actions .race-play{background:var(--mint);color:#001a1f;border-color:transparent}
.race-step{font:800 .8rem 'Space Grotesk';color:var(--gold);text-align:center;letter-spacing:.06em}
.race-match{font-weight:700;margin-top:8px;line-height:1.25}
.race-now{display:grid;grid-template-columns:repeat(3,1fr);gap:7px;margin-top:12px}
.race-now span{background:rgba(0,0,0,.18);border:1px solid rgba(122,252,208,.12);border-radius:10px;padding:8px;font-size:.76rem;color:var(--muted)}
.race-now b{display:block;color:var(--text);font-family:'Space Grotesk';font-size:1rem}
.race-legend{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px;font-size:.76rem;color:var(--muted)}
.race-legend span{display:inline-flex;align-items:center;gap:6px}.race-dot{width:9px;height:9px;border-radius:50%;display:inline-block}
.race-feed{display:grid;gap:9px;margin-top:14px;padding-top:13px;border-top:1px solid rgba(122,252,208,.12)}
.race-feed-title{display:flex;align-items:center;justify-content:space-between;gap:8px;font:800 .68rem 'Space Grotesk';letter-spacing:.08em;
  text-transform:uppercase;color:var(--muted)}
.race-feed-title span{color:var(--gold);letter-spacing:0;text-transform:none}
.race-feed-list{display:grid;gap:6px}
.race-feed-row{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;align-items:center;background:rgba(0,0,0,.16);
  border:1px solid rgba(255,255,255,.06);border-left:3px solid var(--runner,var(--mint));border-radius:9px;padding:7px 8px;min-width:0}
.race-feed-row b{font-size:.8rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.race-feed-row small{font-family:'Space Grotesk';font-weight:800;color:var(--gold);white-space:nowrap}
.race-feed-row .rank-delta{justify-self:end;background:rgba(0,0,0,.18);border-radius:999px;padding:2px 6px;font-size:.68rem}
.race-feed-empty{border:1px dashed rgba(226,246,239,.13);border-radius:9px;padding:8px;color:var(--muted);font-size:.78rem}
.race-board{display:grid;gap:8px;isolation:isolate}
.race-row{display:grid;grid-template-columns:34px minmax(82px,170px) 68px minmax(42px,1fr) 42px 42px;gap:9px;align-items:center;background:rgba(255,255,255,.04);
  border:1px solid rgba(122,252,208,.08);border-left:4px solid var(--runner);border-radius:10px;padding:7px 10px;min-width:0;
  position:relative;will-change:transform,box-shadow;transition:border-color .25s ease,background .25s ease}
.race-row .bar-track{height:10px}.race-fill{height:100%;border-radius:8px;background:linear-gradient(90deg,color-mix(in srgb,var(--runner) 45%,#001a1f),var(--runner));
  transition:width .58s cubic-bezier(.2,.8,.2,1)}
.race-row.leader{border-color:rgba(255,210,122,.45);background:rgba(255,210,122,.06)}
.race-row .rank-delta{justify-self:center;background:rgba(0,0,0,.18);border-radius:999px;padding:3px 7px;font-size:.74rem}
.race-row.is-moving{z-index:5}
.race-row.moved-up{z-index:3;animation:rankGlowUp .76s ease both}.race-row.moved-down{z-index:2;animation:rankGlowDown .76s ease both}
.race-row.moved-up .rank-delta{box-shadow:0 0 0 1px rgba(122,252,208,.25),0 0 18px rgba(122,252,208,.16)}
.race-row.moved-down .rank-delta{box-shadow:0 0 0 1px rgba(255,122,122,.25),0 0 18px rgba(255,122,122,.13)}
@keyframes rankGlowUp{0%,100%{box-shadow:0 0 0 rgba(122,252,208,0)}45%{box-shadow:0 0 0 1px rgba(122,252,208,.28),0 0 24px rgba(122,252,208,.16)}}
@keyframes rankGlowDown{0%,100%{box-shadow:0 0 0 rgba(255,122,122,0)}45%{box-shadow:0 0 0 1px rgba(255,122,122,.24),0 0 24px rgba(255,122,122,.13)}}
.race-round{font-family:'Space Grotesk';font-weight:800;color:var(--runner);text-align:right;font-size:.88rem}
.rank-matrix{display:grid;gap:3px;align-items:center;width:max-content;min-width:100%}
.rank-matrix-name{width:118px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:700;font-size:.78rem}
.rank-matrix-cell{width:31px;height:26px;border-radius:6px;display:flex;align-items:center;justify-content:center;
  font:700 .72rem 'Space Grotesk',sans-serif;color:#001a1f}
.rank-matrix-head{height:24px;color:var(--muted);font:700 .65rem 'Space Grotesk',sans-serif;text-align:center}
.proto-switcher{position:fixed;left:50%;bottom:16px;transform:translateX(-50%);z-index:120;display:flex;align-items:center;gap:10px;
  background:#e2f6ef;color:#001a1f;border:1px solid rgba(255,255,255,.65);border-radius:999px;padding:8px 10px;box-shadow:0 16px 50px rgba(0,0,0,.35)}
.proto-switcher button{border:0;background:#001a1f;color:var(--mint);width:32px;height:32px;border-radius:50%;font:800 1rem 'Space Grotesk';cursor:pointer}
.proto-switcher span{font:800 .78rem 'Space Grotesk';letter-spacing:.04em;white-space:nowrap}
/* PODIUM */
.podium{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;align-items:end;margin-bottom:30px}
.pod{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:20px;text-align:center;position:relative}
.pod .medal{font-size:1.6rem}.pod .nm{font-weight:700;font-size:1.1rem;margin:6px 0 2px}
.pod .sc{font-family:'Space Grotesk';font-size:2rem;font-weight:700;color:var(--mint)}
.pod.p1{transform:translateY(-12px);box-shadow:0 0 0 1px rgba(122,252,208,.25),0 18px 50px rgba(122,252,208,.1)}
/* HEADLINE pair cards */
.duo{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:26px}
.duo .card{display:flex;flex-direction:column;gap:6px}
.duo .names{font-family:'Space Grotesk';font-size:1.5rem;font-weight:700}
.tag{display:inline-block;font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;
  padding:4px 9px;border-radius:999px;background:rgba(122,252,208,.12);color:var(--mint);width:fit-content}
.tag.warm{background:rgba(255,210,122,.14);color:var(--gold)}
.tag.cool{background:rgba(120,180,200,.14);color:#9cd0dd}
/* MATRIX */
.matrix-wrap{overflow-x:auto;padding-bottom:8px}
.matrix{display:inline-grid;gap:2px;margin-top:8px}
.mcell{width:19px;height:19px;border-radius:3px;background:var(--surface2);cursor:default}
.mlabel{font-size:.62rem;color:var(--muted);display:flex;align-items:center}
.mlabel.row{justify-content:flex-end;padding-right:6px;white-space:nowrap}
.mlabel.col{writing-mode:vertical-rl;transform:rotate(180deg);justify-content:flex-end;padding-bottom:6px;height:64px}
.mdiag{background:repeating-linear-gradient(45deg,#063a44,#063a44 3px,#0a4a56 3px,#0a4a56 6px)}
#mtip{position:fixed;pointer-events:none;z-index:90;background:#001a1f;border:1px solid var(--line);
  border-radius:10px;padding:8px 12px;font-size:.82rem;opacity:0;transition:opacity .12s;box-shadow:0 8px 30px rgba(0,0,0,.5)}
.legend{display:flex;align-items:center;gap:10px;margin-top:14px;font-size:.78rem;color:var(--muted)}
.legend .scale{height:10px;width:140px;border-radius:6px;background:linear-gradient(90deg,var(--surface2),var(--mint))}
/* match distribution */
.match-row{padding:14px 0;border-bottom:1px solid var(--line)}
.match-row:last-child{border:0}
.match-top{display:flex;justify-content:space-between;gap:10px;font-weight:600;margin-bottom:8px;font-size:.96rem}
.dist{display:flex;height:24px;border-radius:7px;overflow:hidden;font-size:.72rem;font-weight:700}
.dist span{display:flex;align-items:center;justify-content:center;color:#012;min-width:0;transition:width 1s cubic-bezier(.2,.8,.2,1)}
.dist .s1{background:var(--mint)}.dist .sx{background:#7fa8b2}.dist .s2{background:var(--gold)}
.modal-pill{color:var(--muted);font-size:.84rem;margin-top:6px}
/* chips teams */
.teamchip{display:inline-flex;align-items:center;gap:7px;background:var(--surface);border:1px solid var(--line);
  border-radius:999px;padding:7px 13px;font-size:.9rem;font-weight:600;margin:4px 4px 0 0}
.teamchip .n{color:var(--muted);font-family:'Space Grotesk';font-weight:700}
/* group cards */
.gcard{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:16px}
.gcard .gl{font-family:'Space Grotesk';color:var(--muted);font-size:.8rem;letter-spacing:.1em}
.gcard .fav{font-size:1.15rem;font-weight:700;margin:4px 0 10px;display:flex;align-items:center;gap:8px}
/* fichas */
.search{width:100%;max-width:340px;background:var(--surface);border:1px solid var(--line);color:var(--text);
  border-radius:12px;padding:11px 15px;font-size:.95rem;margin-bottom:20px;font-family:inherit}
.search:focus{outline:none;border-color:var(--mint)}
.ficha{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:18px;transition:.2s}
.ficha:hover{border-color:rgba(122,252,208,.4);transform:translateY(-3px)}
.ficha .fh{display:flex;justify-content:space-between;align-items:start;margin-bottom:10px}
.ficha .fn{font-size:1.2rem;font-weight:700}
.ficha .rk{font-family:'Space Grotesk';font-size:.8rem;color:var(--muted)}
.ficha .lab{font-size:.8rem;margin:2px 0 12px}
.fstats{display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:.84rem}
.fstats .v{font-family:'Space Grotesk';font-weight:700;color:var(--mint)}
.fline{font-size:.82rem;color:var(--muted);margin-top:10px;border-top:1px solid var(--line);padding-top:10px}
.fline b{color:var(--text)}
/* premios */
.award{background:linear-gradient(160deg,var(--surface2),var(--surface));border:1px solid var(--line);
  border-radius:18px;padding:22px;text-align:center}
.award .em{font-size:2.4rem}.award .ti{font-weight:700;margin:8px 0 2px}
.award .wn{font-family:'Space Grotesk';font-size:1.5rem;font-weight:700;color:var(--mint)}
.award .dt{font-size:.82rem;color:var(--muted);margin-top:2px}
/* teaser */
.teaser{text-align:center;padding:46px 22px}
.teaser .em{font-size:3rem}
/* reveal */
.reveal{opacity:0;transform:translateY(20px);transition:opacity .65s ease,transform .65s ease}
.reveal.in{opacity:1;transform:none}
footer{border-top:1px solid var(--line);padding:36px 0 60px;color:var(--muted);font-size:.82rem}
footer .brand{color:var(--mint);height:20px;display:inline-block;vertical-align:middle;margin-right:8px}
footer .brand svg{height:20px;width:auto}
@media(max-width:900px){
  nav.rail{flex-direction:row;width:100%;height:auto;top:auto;bottom:0;padding:8px;overflow-x:auto;gap:2px;
    border-right:0;border-top:1px solid var(--line)}
  nav.rail .brand{display:none}
  nav.rail a{font-size:.7rem;padding:7px 9px}nav.rail a .dot{display:none}
  .nav-group-label{display:none}
  .langtoggle{margin-bottom:0;flex:0 0 auto;order:-1}
  .wrap{padding:0 18px 70px}
  .g3,.g4{grid-template-columns:repeat(2,1fr)}.duo,.podium{grid-template-columns:1fr}
  .chips{grid-template-columns:repeat(2,1fr)}
  .race-layout{grid-template-columns:1fr}
  .proto-switcher{bottom:62px}
}
@media(max-width:560px){
  .g2,.g3,.g4{grid-template-columns:1fr}
  .race-row{grid-template-columns:28px minmax(58px,1fr) 64px 34px 34px;gap:7px}
  .race-row .bar-track{grid-column:2 / -1;grid-row:2}
  .race-round,.race-row .bar-val{font-size:.8rem}
  .race-row .rank-delta{font-size:.68rem;padding:3px 6px}
}
.today-date{font-family:'Space Grotesk';font-size:1.1rem;color:var(--mint);margin-bottom:22px;letter-spacing:.04em}
.today-match{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:24px;margin-bottom:20px}
.today-match .tm-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:8px}
.today-match .tm-teams{font-family:'Space Grotesk';font-size:1.3rem;font-weight:700;display:flex;align-items:center;gap:8px}
.today-match .tm-group{font-size:.78rem;color:var(--muted);background:rgba(122,252,208,.08);padding:4px 10px;border-radius:8px;letter-spacing:.06em}
.today-match .tm-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px}
.today-match .tm-stat{background:rgba(0,0,0,.15);border-radius:12px;padding:12px;text-align:center}
.today-match .tm-stat .val{font-family:'Space Grotesk';font-size:1.3rem;font-weight:700;color:var(--mint)}
.today-match .tm-stat .lab{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:2px}
.today-picks{margin-top:14px}
.today-picks .tp-title{font-size:.82rem;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:10px}
.today-picks-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:6px}
.tp-item{display:flex;align-items:center;gap:6px;padding:6px 10px;background:rgba(0,0,0,.12);border-radius:8px;font-size:.82rem}
.tp-item .tp-name{color:var(--text);font-weight:600;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tp-item .tp-score{color:var(--mint);font-family:'Space Grotesk';font-weight:700;white-space:nowrap}
.recent-meta{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:14px;color:var(--muted);font-size:.86rem}
.score-final{font-family:'Space Grotesk';font-weight:700;font-size:1.7rem;color:var(--gold);white-space:nowrap}
.result-groups{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:18px}
.result-box{background:rgba(0,0,0,.15);border-radius:12px;padding:13px;min-width:0}
.result-box .rb-title{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px}
.result-box .rb-count{font-family:'Space Grotesk';font-weight:700;color:var(--mint)}
.result-box.miss .rb-count{color:var(--red)}
.result-names{display:flex;flex-wrap:wrap;gap:6px}
.result-person{display:inline-flex;gap:5px;align-items:center;background:rgba(255,255,255,.06);border:1px solid var(--line);border-radius:999px;padding:4px 8px;font-size:.78rem;min-width:0}
.result-person b{font-weight:700;max-width:92px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.result-person span{font-family:'Space Grotesk';color:var(--muted);white-space:nowrap}
.no-today{text-align:center;padding:46px 22px}
.no-today .em{font-size:3rem}
.no-today .next{margin-top:22px;text-align:left}
.no-today .next-match{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid var(--line);font-size:.92rem}
.no-today .next-date{color:var(--muted);font-size:.78rem;min-width:80px}
@media(max-width:720px){.result-groups{grid-template-columns:1fr}}
@media(max-width:560px){.today-match .tm-stats{grid-template-columns:1fr}.today-picks-grid{grid-template-columns:repeat(auto-fill,minmax(120px,1fr))}}
.trivia-block{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:16px;padding-top:16px;border-top:1px solid var(--line)}
.trivia-item{background:rgba(0,0,0,.15);border-radius:12px;padding:12px 14px;font-size:.84rem;line-height:1.45}
.trivia-item .trivia-flag{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px}
.trivia-item .trivia-text{color:var(--text)}
@media(max-width:560px){.trivia-block{grid-template-columns:1fr}}
"""

JS = r"""
const D = window.__PORRA__;
const N = D.hero.participants;
const ES2EN = D.es2en || {};
const logo = `__LOGO__`;
let LANG = 'es';
let wrap = null;

function L(es, en){ return LANG === 'es' ? es : en; }
function team(es){ return LANG === 'es' ? es : (ES2EN[es] || es); }
function esc(s){ return String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])); }
function fmt(v){
  let [a,b] = String(v).split('.');
  const th = LANG==='es' ? '.' : ',', dp = LANG==='es' ? ',' : '.';
  a = a.replace(/\B(?=(\d{3})+(?!\d))/g, th);
  return b !== undefined ? a + dp + b : a;
}
function matchdayDateStr(now){
  const d = new Date(now.getTime());
  if (d.getHours() < 6) d.setDate(d.getDate() - 1);
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}
function pf(x){ let s = String(x); if (LANG==='es') s = s.replace('.', ','); return s + '%'; }
function animateCount(node){
  const to = parseFloat(node.dataset.count), dec = parseInt(node.dataset.decimals||'0'), suf = node.dataset.suffix||'';
  const dur = 900, t0 = performance.now();
  (function step(t){ let p = Math.min(1,(t-t0)/dur); p = 1-Math.pow(1-p,3);
    node.textContent = fmt((to*p).toFixed(dec)) + suf; if (p<1) requestAnimationFrame(step); })(t0);
}
function el(t,c,h){ const e = document.createElement(t); if(c) e.className=c; if(h!=null) e.innerHTML=h; return e; }

const NAV_GROUPS = [
  {label:['Jornada','Matchday'], items:[['hoy','Hoy','Today'],['aciertos','Ranking','Ranking'],['ultimos','Últimos','Latest']]},
  {label:['La quiniela','The pool'], items:[['inicio','Resumen','Overview'],['rebeldia','Rebeldía','Maverick'],['gemelas','Afinidad','Affinity'],['estilo','Estilo','Style'],['favoritos','Favoritos','Favourites'],['partidos','Partidos','Matches']]},
  {label:['Participantes','Players'], items:[['lobo','Lobo solitario','Lone wolf'],['fichas','Fichas','Profiles'],['premios','Palmarés','Awards']]},
];
const LABELS = {
  goleador:['🥅 Goleador','🥅 Goal machine'], cerrojo:['🔒 Cerrojo','🔒 Parked bus'],
  empate:['🤝 Amigo del empate','🤝 Draw lover'], rebelde:['🐺 Rebelde','🐺 Maverick'],
  borrego:['🐑 Borrego','🐑 Sheep'], equilibrado:['⚖️ Equilibrado','⚖️ Balanced'],
};
function labelOf(k){ const x = LABELS[k] || LABELS.equilibrado; return L(x[0], x[1]); }

// Live ranking variants, switchable via ?variant=A/B/C. Variant B is the default in #aciertos.
const RANKING_VARIANTS = {
  A: ['Bump chart', 'Bump chart'],
  B: ['Carrera por partido', 'Match-by-match race'],
  C: ['Matriz de posiciones', 'Rank matrix'],
};
const RANK_COLORS = ['#7afcd0','#ffd27a','#ff7a7a','#75e0ff','#b994ff','#ff9d52','#a5e66f','#f56ad0','#8aa0ff','#49d6bb'];
let prototypeKeyListenerReady = false;
let racePlayTimer = null;

function currentRankingVariant(){
  const raw = new URLSearchParams(location.search).get('variant');
  const key = raw ? raw.toUpperCase() : '';
  return RANKING_VARIANTS[key] ? key : 'B';
}
function explicitRankingVariant(){
  const raw = new URLSearchParams(location.search).get('variant');
  const key = raw ? raw.toUpperCase() : '';
  return RANKING_VARIANTS[key] ? key : null;
}
function variantLabel(key){ const x = RANKING_VARIANTS[key] || RANKING_VARIANTS.A; return key + ' · ' + L(x[0], x[1]); }
function liveHistory(){ return (D.live && D.live.progression) || []; }
function latestSnapshot(){ const h = liveHistory(); return h[h.length - 1] || null; }
function matchTitle(m){
  if(!m) return '–';
  return `${m.home_flag} ${esc(team(m.home))} ${m.result.home}-${m.result.away} ${esc(team(m.away))} ${m.away_flag}`;
}
function rankDelta(row){
  if(!row.delta) return `<span class="rank-delta">=</span>`;
  const cls = row.delta > 0 ? 'up' : 'down';
  return `<span class="rank-delta ${cls}">${row.delta > 0 ? '+' : ''}${row.delta}</span>`;
}
function rankDeltaLong(row){
  if(!row.delta) return `<span class="rank-delta">=</span>`;
  const cls = row.delta > 0 ? 'up' : 'down';
  const word = row.delta > 0 ? L('sube','up') : L('baja','down');
  return `<span class="rank-delta ${cls}">${row.delta > 0 ? '+' : ''}${row.delta} ${word}</span>`;
}
function rankingColorMap(){
  const map = {};
  (D.live && D.live.table || []).forEach((r,i) => { map[r.name] = RANK_COLORS[i % RANK_COLORS.length]; });
  return map;
}
function clearRaceTimer(){
  if(racePlayTimer){
    clearInterval(racePlayTimer);
    racePlayTimer = null;
  }
}
function rowMap(rows){
  const map = {};
  rows.forEach(r => { map[r.name] = r; });
  return map;
}
function smoothPath(points){
  if(!points.length) return '';
  let d = `M ${points[0].x.toFixed(1)} ${points[0].y.toFixed(1)}`;
  for(let i=1;i<points.length;i++){
    const a = points[i-1], b = points[i], mid = (a.x + b.x) / 2;
    d += ` C ${mid.toFixed(1)} ${a.y.toFixed(1)} ${mid.toFixed(1)} ${b.y.toFixed(1)} ${b.x.toFixed(1)} ${b.y.toFixed(1)}`;
  }
  return d;
}
function renderRankingState(rows, limit){
  return `<div class="rank-table-compact">${rows.slice(0, limit || rows.length).map(r =>
    `<div class="rank-row">
      <div class="bar-rank">${r.rank}</div>
      <div class="rr-name">${esc(r.name)} ${rankDelta(r)}</div>
      <div class="rr-points">${r.pts}</div>
    </div>`).join('')}</div>`;
}
function rankHeat(rank, total){
  const t = total <= 1 ? 0 : (rank - 1) / (total - 1);
  if(t < .12) return '#ffd27a';
  if(t < .38) return '#7afcd0';
  if(t < .7) return '#75e0ff';
  return '#ff8e7d';
}
function setRankingVariant(next){
  clearRaceTimer();
  const url = new URL(location.href);
  url.searchParams.set('variant', next);
  if(!url.hash) url.hash = '#aciertos';
  history.replaceState(null, '', url);
  rebuild();
  const target = document.getElementById('aciertos');
  if(target) target.scrollIntoView({block:'start'});
}
function cycleRankingVariant(dir){
  const keys = Object.keys(RANKING_VARIANTS);
  const current = currentRankingVariant();
  const idx = keys.indexOf(current);
  setRankingVariant(keys[(idx + dir + keys.length) % keys.length]);
}
function ensurePrototypeKeys(){
  if(prototypeKeyListenerReady) return;
  document.addEventListener('keydown', e => {
    if(!currentRankingVariant()) return;
    const tag = (document.activeElement && document.activeElement.tagName || '').toLowerCase();
    if(tag === 'input' || tag === 'textarea' || (document.activeElement && document.activeElement.isContentEditable)) return;
    if(e.key === 'ArrowLeft'){ e.preventDefault(); cycleRankingVariant(-1); }
    if(e.key === 'ArrowRight'){ e.preventDefault(); cycleRankingVariant(1); }
  });
  prototypeKeyListenerReady = true;
}
function renderPrototypeSwitcher(){
  const old = document.querySelector('.proto-switcher');
  if(old) old.remove();
  const key = explicitRankingVariant();
  if(!key) return;
  ensurePrototypeKeys();
  const bar = el('div','proto-switcher',
    `<button type="button" data-dir="-1" aria-label="${L('Variante anterior','Previous variant')}">‹</button>
     <span>${variantLabel(key)}</span>
     <button type="button" data-dir="1" aria-label="${L('Variante siguiente','Next variant')}">›</button>`);
  bar.addEventListener('click', e => {
    const b = e.target.closest('button[data-dir]');
    if(b) cycleRankingVariant(Number(b.dataset.dir));
  });
  document.body.appendChild(bar);
}
function buildRankingPrototype(s, variant){
  const snap = latestSnapshot();
  if(!snap){
    s.appendChild(el('div','card teaser reveal',
      `<div class="em">📈</div><h3 style="margin:10px 0 6px">${L('Sin histórico todavía','No history yet')}</h3>
       <p class="muted">${L('El ranking necesita al menos un resultado real cargado.','The ranking needs at least one real score loaded.')}</p>`));
    return;
  }
  if(variant === 'B') return buildRankingRaceVariant(s);
  if(variant === 'C') return buildRankingMatrixVariant(s);
  return buildRankingBumpVariant(s);
}
function buildRankingBumpVariant(s){
  const hist = liveHistory(), finalRows = D.live.table, total = finalRows.length;
  const names = finalRows.map(r => r.name);
  const left = 112, right = 166, top = 42, rowH = 31, bottom = 46;
  const plotW = Math.max(740, (hist.length - 1) * 76);
  const width = left + plotW + right, height = top + rowH * Math.max(1, total - 1) + bottom;
  const maps = hist.map(h => rowMap(h.table));
  const xAt = i => left + (hist.length <= 1 ? 0 : (i / (hist.length - 1)) * plotW);
  const yAt = rank => top + (rank - 1) * rowH;
  const grid = hist.map((h,i) => {
    const x = xAt(i);
    const label = i === 0 || i === hist.length - 1 || i % 3 === 0 ? `<text class="bump-axis-label" x="${x}" y="${height-14}" text-anchor="middle">P${h.idx}</text>` : '';
    return `<line class="bump-grid" x1="${x}" y1="${top-18}" x2="${x}" y2="${height-bottom+12}"></line>${label}`;
  }).join('');
  const rankLabels = [1,2,3,5,10,15,20,total].filter((v,i,a) => v <= total && a.indexOf(v) === i).map(r =>
    `<text class="bump-rank-label" x="${left-58}" y="${yAt(r)+4}">#${r}</text>
     <line class="bump-grid" x1="${left-24}" y1="${yAt(r)}" x2="${width-right+24}" y2="${yAt(r)}"></line>`).join('');
  const lines = names.map((name, idx) => {
    const points = maps.map((m,i) => ({x:xAt(i), y:yAt(m[name].rank), row:m[name], snap:hist[i]}));
    const color = RANK_COLORS[idx % RANK_COLORS.length];
    const prominent = idx < 8;
    const first = points[0], last = points[points.length - 1];
    const nodes = prominent ? points.map(p => `<circle class="bump-point" cx="${p.x}" cy="${p.y}" r="4" fill="${color}"><title>${esc(name)} · #${p.row.rank} · ${p.row.pts} pts · ${p.snap.code}</title></circle>`).join('') : '';
    return `<path class="bump-line ${prominent?'':'bump-line-muted'}" d="${smoothPath(points)}" stroke="${color}" stroke-width="${prominent?4:2}"><title>${esc(name)}</title></path>
      ${nodes}
      <text class="bump-name" x="${left-12}" y="${first.y+4}" text-anchor="end" fill="${color}">${esc(name)}</text>
      <text class="bump-name" x="${width-right+12}" y="${last.y+4}" fill="${color}">${esc(name)}</text>
      <text class="bump-score" x="${width-right+112}" y="${last.y+4}" text-anchor="end">${last.row.pts} pts</text>`;
  }).join('');
  const card = el('div','rank-proto-card reveal',
    `<div class="rank-proto-top">
      <div><h3>${L('Evolución puesto a puesto','Position-by-position evolution')}</h3>
      <p class="muted">${L('Cada línea sigue a una persona después de cada partido jugado.','Each line follows one person after every finished match.')}</p></div>
      <div class="rank-proto-meta">
        <span class="rank-state-pill"><b>${D.live.played}</b> ${L('partidos','matches')}</span>
        <span class="rank-state-pill"><b>${esc(finalRows[0].name)}</b> ${L('líder','leader')}</span>
        <span class="rank-state-pill"><b>${finalRows[0].pts}</b> pts</span>
      </div>
    </div>
    <div class="bump-scroll">
      <svg class="bump-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${L('Evolución del ranking partido a partido','Ranking evolution match by match')}">
        ${grid}${rankLabels}${lines}
      </svg>
    </div>
    ${renderRankingState(finalRows, 12)}`);
  s.appendChild(card);
}
function buildRankingRaceVariant(s){
  clearRaceTimer();
  const hist = liveHistory(), finalRows = D.live.table, mx = finalRows[0].pts || 1, colors = rankingColorMap();
  const card = el('div','rank-proto-card reveal',
    `<div class="rank-proto-top">
      <div><h3>${L('Carrera partido a partido','Match-by-match race')}</h3>
      <p class="muted">${L('Barras acumuladas, color por persona y delta de puesto contra el partido anterior.','Accumulated bars, one colour per person, and rank delta against the previous match.')}</p></div>
    </div>
    <div class="race-layout">
      <div class="race-control">
        <div class="k">${L('Partido','Match')}</div>
        <input type="range" min="0" max="${hist.length-1}" value="${hist.length-1}">
        <div class="race-actions">
          <button type="button" class="race-prev" aria-label="${L('Partido anterior','Previous match')}">‹</button>
          <button type="button" class="race-play" aria-label="${L('Reproducir','Play')}">▶</button>
          <div class="race-step"></div>
          <button type="button" class="race-next" aria-label="${L('Partido siguiente','Next match')}">›</button>
        </div>
        <div class="race-match"></div>
        <div class="muted race-date"></div>
        <div class="race-now"></div>
        <div class="race-legend">
          <span><i class="race-dot" style="background:var(--mint)"></i>${L('sube en ranking','rank up')}</span>
          <span><i class="race-dot" style="background:var(--red)"></i>${L('baja en ranking','rank down')}</span>
          <span><i class="race-dot" style="background:var(--gold)"></i>${L('pts del partido','match pts')}</span>
        </div>
        <div class="race-feed"></div>
      </div>
      <div class="race-board"></div>
    </div>`);
  const input = card.querySelector('input');
  const board = card.querySelector('.race-board');
  const match = card.querySelector('.race-match');
  const date = card.querySelector('.race-date');
  const step = card.querySelector('.race-step');
  const now = card.querySelector('.race-now');
  const feed = card.querySelector('.race-feed');
  const play = card.querySelector('.race-play');
  const prev = card.querySelector('.race-prev');
  const next = card.querySelector('.race-next');
  function setPlaying(isPlaying){
    play.textContent = isPlaying ? 'Ⅱ' : '▶';
    play.setAttribute('aria-label', isPlaying ? L('Pausar','Pause') : L('Reproducir','Play'));
  }
  function stopPlaying(){
    clearRaceTimer();
    setPlaying(false);
  }
  function goTo(idx){
    input.value = String(Math.max(0, Math.min(hist.length - 1, idx)));
    paint();
  }
  function collectRowRects(){
    const rects = new Map();
    board.querySelectorAll('.race-row').forEach(row => {
      rects.set(row.dataset.name, row.getBoundingClientRect());
    });
    return rects;
  }
  function rowHtml(r,i){
    return `<div class="race-row ${i===0?'leader':''} ${r.delta>0?'moved-up':(r.delta<0?'moved-down':'')}" data-name="${esc(r.name)}" style="--runner:${colors[r.name] || RANK_COLORS[i % RANK_COLORS.length]}">
      <div class="bar-rank">${r.rank}</div>
      <div class="bar-name">${esc(r.name)}</div>
      ${rankDeltaLong(r)}
      <div class="bar-track"><div class="race-fill" style="width:${(r.pts/mx*100).toFixed(1)}%"></div></div>
      <div class="bar-val">${r.pts}</div>
      <div class="race-round">+${r.round_pts || 0}</div>
    </div>`;
  }
  function feedRow(row, valueHtml){
    return `<div class="race-feed-row" style="--runner:${colors[row.name] || 'var(--mint)'}">
      <b>${esc(row.name)}</b>
      ${valueHtml}
    </div>`;
  }
  function raceFeedHtml(snap){
    const movers = snap.table
      .filter(r => r.delta !== 0)
      .sort((a,b) => Math.abs(b.delta) - Math.abs(a.delta) || b.round_pts - a.round_pts || a.rank - b.rank)
      .slice(0, 4);
    const scorers = snap.table
      .filter(r => (r.round_pts || 0) > 0)
      .sort((a,b) => b.round_pts - a.round_pts || a.rank - b.rank)
      .slice(0, 5);
    const movementRows = movers.length
      ? movers.map(r => feedRow(r, rankDeltaLong(r))).join('')
      : `<div class="race-feed-empty">${L('Sin cambios de puesto en este partido.','No rank changes in this match.')}</div>`;
    const scorerRows = scorers.length
      ? scorers.map(r => feedRow(r, `<small>+${r.round_pts} pts</small>`)).join('')
      : `<div class="race-feed-empty">${L('Sin puntos repartidos todavía.','No points awarded yet.')}</div>`;
    return `<div class="race-feed-title">${L('Movimientos','Movements')}<span>${movers.length}</span></div>
      <div class="race-feed-list">${movementRows}</div>
      <div class="race-feed-title">${L('Puntos del partido','Match points')}</div>
      <div class="race-feed-list">${scorerRows}</div>`;
  }
  function animateRowsFrom(previousRects){
    if(!previousRects || !previousRects.size || !board.isConnected || typeof Element === 'undefined') return;
    const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if(reduceMotion) return;
    board.querySelectorAll('.race-row').forEach(row => {
      const before = previousRects.get(row.dataset.name);
      if(!before || typeof row.animate !== 'function') return;
      const after = row.getBoundingClientRect();
      const dx = before.left - after.left;
      const dy = before.top - after.top;
      if(Math.abs(dx) < .5 && Math.abs(dy) < .5) return;
      const duration = Math.min(880, Math.max(540, Math.round(430 + Math.abs(dy) * 1.05)));
      row.classList.add('is-moving');
      const animation = row.animate([
        {transform:`translate3d(${dx}px,${dy}px,0)`},
        {transform:'translate3d(0,0,0)'}
      ], {duration, easing:'cubic-bezier(.16,.9,.2,1)', fill:'both'});
      const cleanup = () => {
        row.classList.remove('is-moving');
        animation.cancel();
      };
      animation.finished.then(cleanup).catch(() => row.classList.remove('is-moving'));
    });
  }
  function paint(options){
    const animate = !options || options.animate !== false;
    const previousRects = animate ? collectRowRects() : null;
    const snap = hist[Number(input.value)];
    const leader = snap.table[0];
    const movers = snap.table.filter(r => r.delta !== 0).length;
    const pointsNow = snap.table.reduce((sum,r) => sum + (r.round_pts || 0), 0);
    step.textContent = `${snap.idx}/${hist.length}`;
    match.innerHTML = `${snap.idx}. ${matchTitle(snap)}`;
    date.textContent = `${snap.date} · ${L('Grupo','Group')} ${snap.group}`;
    now.innerHTML = `
      <span><b>${esc(leader.name)}</b>${L('líder','leader')} · ${leader.pts} pts</span>
      <span><b>${movers}</b>${L('movimientos','moves')}</span>
      <span><b>+${pointsNow}</b>${L('pts repartidos','pts awarded')}</span>`;
    feed.innerHTML = raceFeedHtml(snap);
    board.innerHTML = snap.table.map(rowHtml).join('');
    requestAnimationFrame(() => animateRowsFrom(previousRects));
  }
  input.addEventListener('input', () => { stopPlaying(); paint(); });
  prev.addEventListener('click', () => { stopPlaying(); goTo(Number(input.value) - 1); });
  next.addEventListener('click', () => { stopPlaying(); goTo(Number(input.value) + 1); });
  play.addEventListener('click', () => {
    if(racePlayTimer){ stopPlaying(); return; }
    if(Number(input.value) >= hist.length - 1) goTo(0);
    setPlaying(true);
    racePlayTimer = setInterval(() => {
      const idx = Number(input.value);
      if(idx >= hist.length - 1){ stopPlaying(); return; }
      goTo(idx + 1);
    }, 900);
  });
  paint({animate:false});
  s.appendChild(card);
}
function buildRankingMatrixVariant(s){
  const hist = liveHistory(), finalRows = D.live.table, total = finalRows.length;
  const cols = `118px repeat(${hist.length},31px)`;
  const header = `<div></div>` + hist.map(h => `<div class="rank-matrix-head">P${h.idx}</div>`).join('');
  const maps = hist.map(h => rowMap(h.table));
  const rows = finalRows.map(person => {
    const cells = hist.map((h,i) => {
      const r = maps[i][person.name];
      return `<div class="rank-matrix-cell" style="background:${rankHeat(r.rank,total)}" title="${esc(person.name)} · ${h.code} · #${r.rank} · ${r.pts} pts">${r.rank}</div>`;
    }).join('');
    return `<div class="rank-matrix-name">${esc(person.name)}</div>${cells}`;
  }).join('');
  const card = el('div','rank-proto-card reveal',
    `<div class="rank-proto-top">
      <div><h3>${L('Todas las posiciones, sin cruces','Every rank without crossing lines')}</h3>
      <p class="muted">${L('Cada celda es el puesto de una persona tras ese partido; dorado arriba, rojo abajo.','Each cell is a person’s rank after that match; gold is top, red is bottom.')}</p></div>
      <div class="rank-proto-meta">
        <span class="rank-state-pill"><b>${total}</b> ${L('personas','people')}</span>
        <span class="rank-state-pill"><b>${hist.length}</b> ${L('partidos','matches')}</span>
      </div>
    </div>
    <div class="rank-matrix-scroll">
      <div class="rank-matrix" style="grid-template-columns:${cols}">${header}${rows}</div>
    </div>
    ${renderRankingState(finalRows, total)}`);
  s.appendChild(card);
}

/* ---- NAV (built once, labels re-rendered on language change) ---- */
const rail = el('nav','rail');
document.body.appendChild(rail);
function renderRail(){
  const navHtml = NAV_GROUPS.map(g =>
    `<div class="nav-group-label">${L(g.label[0],g.label[1])}</div>` +
    g.items.map(([id,es,en]) => `<a href="#${id}" data-id="${id}"><span class="dot"></span>${L(es,en)}</a>`).join('')
  ).join('');
  rail.innerHTML = `<a class="brand" href="#hoy">${logo}</a>` +
    `<div class="langtoggle"><button data-l="es">ES</button><button data-l="en">EN</button></div>` +
    navHtml;
  const on = rail.querySelector('[data-l="'+LANG+'"]'); if (on) on.classList.add('on');
}
rail.addEventListener('click', e => {
  const b = e.target.closest('button[data-l]');
  if (b && b.dataset.l !== LANG){ LANG = b.dataset.l; rebuild(); }
});

/* shared tooltip for the matrix */
const tip = el('div'); tip.id = 'mtip'; document.body.appendChild(tip);

function section(id, kicker, title, sub){
  const s = el('section','sec'); s.id = id;
  s.appendChild(el('div','sec-head reveal',
    `<div class="kicker">${kicker}</div><h2>${title}</h2>${sub?`<p class="sub">${sub}</p>`:''}`));
  wrap.appendChild(s); return s;
}

/* ---- HERO ---- */
function chip(val, lab, dec){
  return `<div class="chip reveal"><div class="big" data-count="${val}"${dec?` data-decimals="${dec}"`:''}>0</div><div class="lab">${lab}</div></div>`;
}
function chipText(val, lab){
  return `<div class="chip reveal"><div class="big">${val}</div><div class="lab">${lab}</div></div>`;
}
function buildHero(){
  const h = el('header','hero'); h.id = 'inicio';
  const ts = D.hero.top_scoreline, bp = D.best_pair;
  h.innerHTML = `
    <div class="eyebrow reveal">${L('Reveni · Porra Mundial 2026','Reveni · World Cup 2026 Pool')}</div>
    <h1 class="reveal">${L('La quiniela,<br>bajo el microscopio.','The pool,<br>under the microscope.')}</h1>
    <p class="lead reveal">${L(
      'Hemos analizado las '+N+' quinielas de la porra del Mundial: quién se la juega, quién va a lo seguro, quién piensa igual que quién y qué partidos nos parten en dos.',
      'We crunched all '+N+' predictions in the office World Cup pool: who gambles, who plays it safe, who thinks like who, and which matches split us in two.')}</p>
    <div class="chips">
      ${chip(N, L('Participantes','Players'))}
      ${chip(D.hero.matches, L('Partidos de grupos','Group matches'))}
      ${chip(D.hero.groups, L('Grupos','Groups'))}
      ${chip(N*D.hero.matches, L('Marcadores analizados','Predictions analysed'))}
      ${chip(D.hero.total_goals, L('Goles apostados','Goals predicted'))}
      ${chip(D.hero.avg_goals_match, L('Goles/partido de media','Avg goals per match'), 1)}
      ${chipText(ts?ts.score:'–', L('Marcador estrella','Star scoreline')+' ('+pf(ts?ts.pct:0)+')')}
      ${chipText(pf(bp.pct), L('Máx. afinidad','Top affinity')+' · '+esc(bp.a)+' & '+esc(bp.b))}
    </div>
    <div class="scrollcue">${L('desliza para empezar','scroll to start')} ↓</div>`;
  wrap.appendChild(h);
}

/* ---- HOY ---- */
function buildHoy(){
  const s = section('hoy', L('⚽ Hoy','⚽ Today'),
    L('Los partidos de hoy','Today\'s matches'),
    L('Qué se juega hoy y qué ha puesto cada uno.','What\'s on today and what everyone predicted.'));
  const allM = D.today.matches || [];
  const todayStr = matchdayDateStr(new Date());
  const todayMatches = allM.filter(m => m.date === todayStr);
  if(!todayMatches.length){
    const upcoming = allM.filter(m => m.date > todayStr).sort((a,b) => a.date.localeCompare(b.date)).slice(0,6);
    let nextHtml = '';
    if(upcoming.length){
      const grouped = {};
      upcoming.forEach(m => { if(!grouped[m.date]) grouped[m.date] = []; grouped[m.date].push(m); });
      nextHtml = '<div class="next">' + Object.entries(grouped).map(([dt, ms]) =>
        ms.map(m => `<div class="next-match"><span class="next-date">${dt.slice(5)}</span>${m.home_flag} ${esc(team(m.home))} – ${esc(team(m.away))} ${m.away_flag}</div>`).join('')
      ).join('') + '</div>';
    }
    s.appendChild(el('div','card no-today reveal',
      `<div class="em">📅</div><h3 style="margin:10px 0 6px">${L('Hoy no hay partidos','No matches today')}</h3>
       <p class="muted">${L('Próximos partidos:','Upcoming matches:')}</p>${nextHtml}`));
    return;
  }
  const dateObj = new Date(todayStr + 'T12:00:00');
  const opts = {weekday:'long', day:'numeric', month:'long', year:'numeric'};
  const dateStr = dateObj.toLocaleDateString(LANG==='es'?'es-ES':'en-US', opts);
  s.appendChild(el('div','today-date reveal', dateStr.charAt(0).toUpperCase() + dateStr.slice(1)));
  todayMatches.forEach(m => {
    const o = m.outcome_dist, tot = (o['1']||0)+(o['X']||0)+(o['2']||0);
    const pct = k => tot ? Math.round((o[k]||0)/tot*100) : 0;
    const uniqueHtml = m.most_unique_pick
      ? `<span class="muted">${L('🔥 El más atrevido:','🔥 Boldest call:')}</span> <b>${esc(m.most_unique_pick.name)}</b> <span class="mint">${m.most_unique_pick.score}</span>`
      : '';
    const picksHtml = m.picks.map(p =>
      `<div class="tp-item"><span class="tp-name">${esc(p.name)}</span><span class="tp-score">${p.home!=null?p.home+'-'+p.away:'–'}</span></div>`
    ).join('');
    s.appendChild(el('div','today-match reveal',
      `<div class="tm-head">
        <div class="tm-teams">${m.home_flag} ${esc(team(m.home))} – ${esc(team(m.away))} ${m.away_flag}</div>
        <span class="tm-group">${L('GRUPO','GROUP')} ${m.group}</span>
      </div>
      <div class="tm-stats">
        <div class="tm-stat"><div class="val">${pct('1')}%</div><div class="lab">${L('Gana','Win')} ${esc(team(m.home))}</div></div>
        <div class="tm-stat"><div class="val">${pct('X')}%</div><div class="lab">${L('Empate','Draw')}</div></div>
        <div class="tm-stat"><div class="val">${pct('2')}%</div><div class="lab">${L('Gana','Win')} ${esc(team(m.away))}</div></div>
      </div>
      <div class="tm-stats" style="grid-template-columns:1fr 1fr">
        <div class="tm-stat"><div class="val">${m.modal_scoreline}</div><div class="lab">${L('Marcador más repetido','Most common scoreline')} (${pf(Math.round(m.modal_scoreline_share*100))})</div></div>
        <div class="tm-stat"><div class="val" style="font-size:1rem">${uniqueHtml||'–'}</div><div class="lab">${L('Pick único más salvaje','Wildest unique pick')}</div></div>
      </div>
      <div class="today-picks">
        <div class="tp-title">${L('Qué ha puesto cada uno','What everyone picked')}</div>
        <div class="today-picks-grid">${picksHtml}</div>
      </div>
      <div class="trivia-block">
        <div class="trivia-item">
          <div class="trivia-flag">${L('🤯 ¿Sabías que…?','🤯 Did you know…?')} ${m.home_flag} ${esc(team(m.home))}</div>
          <div class="trivia-text">${esc(L(m.home_trivia.es, m.home_trivia.en))}</div>
        </div>
        <div class="trivia-item">
          <div class="trivia-flag">${L('🤯 ¿Sabías que…?','🤯 Did you know…?')} ${m.away_flag} ${esc(team(m.away))}</div>
          <div class="trivia-text">${esc(L(m.away_trivia.es, m.away_trivia.en))}</div>
        </div>
      </div>`));
  });
}

/* ---- ÚLTIMOS RESULTADOS ---- */
function buildUltimos(){
  const s = section('ultimos', L('🧾 Últimos','🧾 Latest'),
    L('Últimos partidos jugados','Latest finished matches'),
    L('Resultado final y reparto de alegrías: plenos, signos acertados y los que han palmado.',
      'Final score and who got it right: exact hits, correct outcomes and misses.'));
  const recent = (D.recent_results && D.recent_results.matches) || [];
  if(!recent.length){
    s.appendChild(el('div','card teaser reveal',
      `<div class="em">🧾</div><h3 style="margin:10px 0 6px">${L('Aún no hay resultados cargados','No results loaded yet')}</h3>
       <p class="muted">${L('Cuando rellenes marcadores en <b>Real results</b>, aquí aparecerán los últimos partidos terminados.','Once scores are filled in <b>Real results</b>, the latest finished matches will appear here.')}</p>`));
    return;
  }
  s.appendChild(el('div','muted reveal', `${recent.length} ${L('últimos de','latest of')} ${D.recent_results.total} ${L('partidos jugados','finished matches')}`));
  recent.forEach(m => {
    const people = arr => arr.length
      ? arr.map(p => `<span class="result-person"><b>${esc(p.name)}</b><span>${esc(p.pick)}</span></span>`).join('')
      : `<span class="muted">${L('Nadie','Nobody')}</span>`;
    const dateObj = new Date(m.date + 'T12:00:00');
    const dateStr = dateObj.toLocaleDateString(LANG==='es'?'es-ES':'en-US', {day:'numeric', month:'short'});
    s.appendChild(el('div','today-match reveal',
      `<div class="tm-head">
        <div class="tm-teams">${m.home_flag} ${esc(team(m.home))} – ${esc(team(m.away))} ${m.away_flag}</div>
        <span class="score-final">${m.result.home}-${m.result.away}</span>
      </div>
      <div class="recent-meta"><span>${dateStr}</span><span>${L('Grupo','Group')} ${m.group}</span><span>${L('Resultado final','Final score')}</span></div>
      <div class="result-groups">
        <div class="result-box">
          <div class="rb-title">${L('Pleno','Exact score')} <span class="rb-count">${m.exact.length}</span></div>
          <div class="result-names">${people(m.exact)}</div>
        </div>
        <div class="result-box">
          <div class="rb-title">${L('Signo','Outcome')} <span class="rb-count">${m.sign.length}</span></div>
          <div class="result-names">${people(m.sign)}</div>
        </div>
        <div class="result-box miss">
          <div class="rb-title">${L('Palmada','Missed')} <span class="rb-count">${m.miss.length}</span></div>
          <div class="result-names">${people(m.miss)}</div>
        </div>
      </div>`));
  });
}

/* ---- REBELDÍA ---- */
function buildRebeldia(){
  const s = section('rebeldia', L('01 · La estrella','01 · The star'),
    L('El ranking de rebeldía 🐺','The maverick ranking 🐺'),
    L('Combina dos cosas en cada partido: cuántas veces eliges el signo (1/X/2) en minoría, y cuánto se aleja tu marcador del marcador típico del grupo. 100 = el más inconformista. 0 = el más borrego.',
      'Two things per match: how often you back the (1/X/2) outcome in the minority, and how far your scoreline sits from the group typical one. 100 = the biggest maverick. 0 = the biggest sheep.'));
  const R = D.rebeldia, max = R[0].index || 100;
  const pod = el('div','podium reveal');
  const order = [1,0,2], medals = ['🥈','🥇','🥉'];
  order.forEach((oi,k) => { const r = R[oi]; if(!r) return;
    pod.appendChild(el('div','pod '+(oi===0?'p1':''),
      `<div class="medal">${medals[k]}</div><div class="nm">${esc(r.name)}</div>
       <div class="sc" data-count="${r.index}" data-decimals="1">0</div><div class="k">${L('índice','index')}</div>`)); });
  s.appendChild(pod);
  const list = el('div','reveal');
  R.forEach((r,i) => {
    const badge = i===0?'🐺 ':(i===R.length-1?'🐑 ':'');
    const cls = i===R.length-1?'gold':'';
    list.appendChild(el('div','bar-row',
      `<div class="bar-rank">${i+1}</div><div class="bar-name">${badge}${esc(r.name)}</div>
       <div class="bar-track"><div class="bar-fill ${cls}" data-w="${(r.index/max*100).toFixed(1)}"></div></div>
       <div class="bar-val" data-count="${r.index}" data-decimals="1">0</div>`));
  });
  s.appendChild(list);
}

/* ---- AFINIDAD ---- */
function buildAfinidad(){
  const s = section('gemelas', L('02 · Afinidad','02 · Affinity'),
    L('¿Quién piensa como quién? 💞','Who thinks like who? 💞'),
    L('Comparamos las quinielas marcador a marcador. Cuanto más brillante el cuadro, más coinciden esas dos personas en el resultado exacto.',
      'We compare predictions scoreline by scoreline. The brighter the cell, the more those two people match on the exact result.'));
  const bp = D.best_pair, wp = D.worst_pair;
  const duo = el('div','duo reveal');
  duo.innerHTML = `
    <div class="card glow"><span class="tag">${L('💞 Almas gemelas','💞 Soulmates')}</span>
      <div class="names">${esc(bp.a)} & ${esc(bp.b)}</div>
      <div class="muted">${bp.pct}% ${L('de marcadores','of scorelines')} <b class="mint">${L('idénticos','identical')}</b>${bp.pct>=100?L(' — sí, los 72. 👀 ¿alguien ha copiado?',' — yes, all 72. 👀 did someone copy?'):''}</div></div>
    <div class="card"><span class="tag cool">${L('🧊 Polos opuestos','🧊 Opposites')}</span>
      <div class="names">${esc(wp.a)} & ${esc(wp.b)}</div>
      <div class="muted">${L('los que más se contradicen (distancia media '+wp.dist+' goles por partido)','the most contradictory pair (avg distance '+wp.dist+' goals per match)')}</div></div>`;
  s.appendChild(duo);
  const names = D.matrix.names, sim = D.matrix.sim;
  let mx = 0; for(let i=0;i<N;i++) for(let j=0;j<N;j++) if(i!==j) mx = Math.max(mx, sim[i][j]);
  const wrapm = el('div','matrix-wrap reveal');
  const grid = el('div','matrix');
  grid.style.gridTemplateColumns = `120px repeat(${N},19px)`;
  grid.appendChild(el('div','mlabel'));
  names.forEach(n => grid.appendChild(el('div','mlabel col', esc(n))));
  for(let i=0;i<N;i++){
    grid.appendChild(el('div','mlabel row', esc(names[i])));
    for(let j=0;j<N;j++){
      if(i===j){ grid.appendChild(el('div','mcell mdiag')); continue; }
      const v = sim[i][j], a = mx ? Math.max(.05, v/mx) : 0;
      const c = el('div','mcell');
      c.style.background = `rgba(122,252,208,${a.toFixed(3)})`;
      c.dataset.t = `${esc(names[i])} ↔ ${esc(names[j])} · ${v}% ${L('idénticos','identical')}`;
      grid.appendChild(c);
    }
  }
  wrapm.appendChild(grid);
  wrapm.appendChild(el('div','legend',`<span>${L('menos afín','less alike')}</span><span class="scale"></span><span>${L('más afín','more alike')}</span>`));
  s.appendChild(wrapm);
  grid.addEventListener('mouseover', e => { const t = e.target.dataset.t; if(t){ tip.textContent = t; tip.style.opacity = 1; } });
  grid.addEventListener('mousemove', e => { tip.style.left = (e.clientX+14)+'px'; tip.style.top = (e.clientY+14)+'px'; });
  grid.addEventListener('mouseout', () => { tip.style.opacity = 0; });
}

/* ---- ESTILO ---- */
function buildEstilo(){
  const s = section('estilo', L('03 · Estilo','03 · Style'),
    L('Goleadores vs cerrojos 🥅','Goal-fests vs parked buses 🥅'),
    L('Media de goles que cada uno mete en sus marcadores, qué marcadores se repiten más en toda la porra y quién es el rey del empate.',
      'Average goals in everyone scorelines, the most repeated results across the pool, and who loves a draw.'));
  const st = D.style.slice().sort((a,b) => b.avg_goals-a.avg_goals);
  const gol = st[0], cer = st[st.length-1];
  const emp = D.style.slice().sort((a,b) => b.pct_draws-a.pct_draws)[0];
  const duo = el('div','grid g3 reveal');
  duo.innerHTML = `
    <div class="card glow"><span class="k">${L('🥅 El goleador','🥅 The goal machine')}</span><h3>${esc(gol.name)}</h3>
      <div class="big-num">${gol.avg_goals}</div><div class="muted">${L('goles de media por partido','avg goals per match')}</div></div>
    <div class="card"><span class="k">${L('🔒 El cerrojo','🔒 The parked bus')}</span><h3>${esc(cer.name)}</h3>
      <div class="big-num">${cer.avg_goals}</div><div class="muted">${L('el más tacaño con los goles','the stingiest with goals')}</div></div>
    <div class="card"><span class="k">${L('🤝 Amigo del empate','🤝 Draw lover')}</span><h3>${esc(emp.name)}</h3>
      <div class="big-num">${pf(emp.pct_draws)}</div><div class="muted">${L('de sus partidos, en tablas','of their matches drawn')}</div></div>`;
  s.appendChild(duo);
  const list = el('div','reveal'); list.style.marginTop = '22px';
  const mxg = st[0].avg_goals;
  list.appendChild(el('div','k', L('Media de goles por quiniela','Average goals per prediction')));
  st.forEach((r,i) => list.appendChild(el('div','bar-row',
    `<div class="bar-rank">${i+1}</div><div class="bar-name">${esc(r.name)}</div>
     <div class="bar-track"><div class="bar-fill" data-w="${(r.avg_goals/mxg*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${r.avg_goals}" data-decimals="2">0</div>`)));
  s.appendChild(list);
  const cs = el('div','card reveal'); cs.style.marginTop = '22px';
  cs.innerHTML = `<span class="k">${L('Marcadores más apostados (de las '+fmt(N*D.hero.matches)+' apuestas)','Most-predicted scorelines (out of '+fmt(N*D.hero.matches)+' bets)')}</span>`;
  const mxc = D.common_scorelines[0].count;
  D.common_scorelines.forEach(c => cs.appendChild(el('div','bar-row',
    `<div class="bar-rank">${c.score}</div><div class="bar-name muted">${pf(c.pct)}</div>
     <div class="bar-track"><div class="bar-fill" data-w="${(c.count/mxc*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${c.count}">0</div>`)));
  s.appendChild(cs);
}

/* ---- FAVORITOS ---- */
function buildFavoritos(){
  const s = section('favoritos', L('04 · Favoritos','04 · Favourites'),
    L('Los favoritos del torneo 🏆','Tournament favourites 🏆'),
    L('El consenso de la oficina: a quién ve todo el mundo ganando su grupo, qué selecciones se respetan más y qué grupo nos tiene más divididos.',
      'The office consensus: who everyone sees topping their group, the most respected teams, and the group that divides us most.'));
  const gc = D.group_consensus.slice().sort((a,b) => b.agreement-a.agreement);
  const grid = el('div','grid g4 reveal');
  gc.forEach(g => grid.appendChild(el('div','gcard',
    `<div class="gl">${L('GRUPO','GROUP')} ${g.group}</div>
     <div class="fav">${g.flag} ${esc(team(g.favorite))}</div>
     <div class="bar-track"><div class="bar-fill" data-w="${g.agreement}"></div></div>
     <div class="muted" style="font-size:.8rem;margin-top:6px">${pf(g.agreement)} ${L('lo ve primero','see them top')}</div>`)));
  s.appendChild(grid);
  const mr = el('div','card reveal'); mr.style.marginTop = '22px';
  mr.innerHTML = `<span class="k">${L('Selecciones más respetadas (prestigio: 1.º=3, 2.º=2, 3.º=1)','Most respected teams (prestige: 1st=3, 2nd=2, 3rd=1)')}</span>`;
  const mxr = D.most_respected[0].count;
  D.most_respected.forEach((t,i) => mr.appendChild(el('div','bar-row',
    `<div class="bar-rank">${i+1}</div><div class="bar-name">${t.flag} ${esc(team(t.team))}</div>
     <div class="bar-track"><div class="bar-fill gold" data-w="${(t.count/mxr*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${t.count}">0</div>`)));
  s.appendChild(mr);
  const two = el('div','grid g2 reveal'); two.style.marginTop = '22px';
  const pol = D.polarizing.slice(0,4).map(p => {
    const top = Object.entries(p.dist).slice(0,3).map(([t,c]) => `${esc(team(t))} (${c})`).join(' · ');
    return `<div style="padding:8px 0;border-bottom:1px solid var(--line)"><b>${L('Grupo','Group')} ${p.group}</b> — ${pf(p.agreement)} ${L('de acuerdo','agreement')}<br><span class="muted" style="font-size:.84rem">${top}</span></div>`;
  }).join('');
  const th = D.top_thirds.slice(0,10).map(t => `<span class="teamchip">${t.flag} ${esc(team(t.team))} <span class="n">${t.count}</span></span>`).join('');
  two.innerHTML = `
    <div class="card"><span class="k">${L('🔥 Los grupos que más nos dividen','🔥 The groups that divide us most')}</span><div style="margin-top:10px">${pol}</div></div>
    <div class="card"><span class="k">${L('🎟️ Terceros más elegidos para colarse','🎟️ Most-picked third places')}</span><div style="margin-top:10px">${th}</div></div>`;
  s.appendChild(two);
}

/* ---- PARTIDOS ---- */
function buildPartidos(){
  const s = section('partidos', L('05 · Partidos','05 · Matches'),
    L('Los partidos que nos parten en dos 🔪','The matches that split us 🔪'),
    L('Dónde hay guerra de pronósticos (poco acuerdo en el 1/X/2) y dónde casi todos ponemos lo mismo.',
      'Where there is a prediction war (little agreement on 1/X/2) and where we nearly all agree.'));
  function matchBlock(m){
    const o = m.outcome_dist, tot = (o['1']||0)+(o['X']||0)+(o['2']||0);
    const p = k => tot ? Math.round((o[k]||0)/tot*100) : 0;
    return `<div class="match-row">
      <div class="match-top"><span>${m.home_flag} ${esc(team(m.home))} – ${esc(team(m.away))} ${m.away_flag}</span>
        <span class="muted">${pf(Math.round(m.outcome_agreement*100))} ${L('acuerdo','agree')}</span></div>
      <div class="dist">
        <span class="s1" style="width:${p('1')}%">${p('1')?p('1')+'%':''}</span>
        <span class="sx" style="width:${p('X')}%">${p('X')?'X '+p('X')+'%':''}</span>
        <span class="s2" style="width:${p('2')}%">${p('2')?p('2')+'%':''}</span></div>
      <div class="modal-pill">${L('marcador más repetido','most common scoreline')}: <b class="mint">${m.modal_scoreline}</b> (${pf(Math.round(m.modal_scoreline_share*100))})</div>
    </div>`;
  }
  const two = el('div','grid g2 reveal');
  two.innerHTML = `
    <div class="card"><span class="k">${L('😵‍💫 Máxima guerra de pronósticos','😵‍💫 Maximum disagreement')}</span><div style="margin-top:8px">${D.divisive.map(matchBlock).join('')}</div></div>
    <div class="card"><span class="k">${L('🤖 Casi unanimidad','🤖 Near unanimity')}</span><div style="margin-top:8px">${D.unanimous.map(matchBlock).join('')}</div></div>`;
  s.appendChild(two);
}

/* ---- LOBO SOLITARIO ---- */
function buildLobo(){
  const s = section('lobo', L('06 · Atrevimiento','06 · Boldness'),
    L('Lobo solitario 🐺','Lone wolf 🐺'),
    L('Marcadores que solo apostó UNA persona en todo el grupo. El que más tiene, es el que más se aleja del rebaño.',
      'Scorelines only ONE person predicted in the whole pool. The more you have, the further from the flock you roam.'));
  const mx = D.lobo[0].count;
  const list = el('div','reveal');
  D.lobo.forEach((r,i) => list.appendChild(el('div','bar-row',
    `<div class="bar-rank">${i+1}</div><div class="bar-name">${i===0?'🐺 ':''}${esc(r.name)}</div>
     <div class="bar-track"><div class="bar-fill" data-w="${(r.count/mx*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${r.count}">0</div>`)));
  s.appendChild(list);
  const ex = el('div','card reveal'); ex.style.marginTop = '22px';
  ex.innerHTML = `<span class="k">${L('Apuestas en solitario más salvajes','Wildest solo calls')}</span><div style="margin-top:10px">` +
    D.unique_examples.map(e => `<div style="padding:7px 0;border-bottom:1px solid var(--line)">
      <b>${esc(e.name)}</b> <span class="muted">${L('es el único que firmó','is the only one who called')}</span> ${e.flags} <b class="mint">${esc(team(e.home))} ${e.score} ${esc(team(e.away))}</b></div>`).join('') + '</div>';
  s.appendChild(ex);
}

/* ---- FICHAS ---- */
function buildFichas(){
  const s = section('fichas', L('07 · Uno a uno','07 · One by one'),
    L('La ficha de cada uno 🪪',"Everyone's card 🪪"),
    L('Resumen por persona: su puesto en rebeldía, su estilo, su gemelo y su apuesta más loca. Busca tu nombre.',
      'A summary per person: their maverick rank, style, twin and wildest bet. Search your name.'));
  const inp = el('input','search'); inp.placeholder = L('🔎 Busca tu nombre…','🔎 Search your name…');
  s.appendChild(inp);
  const grid = el('div','grid g3 reveal');
  D.cards.slice().sort((a,b) => a.rebel_rank-b.rebel_rank).forEach(c => {
    const big = c.biggest ? `${c.biggest.flags} ${esc(team(c.biggest.home))} ${c.biggest.score} ${esc(team(c.biggest.away))}` : '–';
    const card = el('div','ficha'); card.dataset.name = c.name.toLowerCase();
    card.innerHTML = `
      <div class="fh"><div><div class="fn">${esc(c.name)}</div><div class="lab">${labelOf(c.label)}</div></div>
        <div class="rk">#${c.rebel_rank}<br>${L('rebeldía','maverick')}</div></div>
      <div class="fstats">
        <div>${L('Índice rebeldía','Maverick index')}<br><span class="v">${c.rebel_index}</span></div>
        <div>${L('Goles/partido','Goals/match')}<br><span class="v">${c.avg_goals}</span></div>
        <div>${L('% empates','% draws')}<br><span class="v">${pf(c.pct_draws)}</span></div>
        <div>${L('Apuestas únicas','Unique bets')}<br><span class="v">${c.lobo}</span></div>
      </div>
      <div class="fline">${L('👯 Gemelo:','👯 Twin:')} <b>${esc(c.twin)}</b> (${pf(c.twin_pct)})</div>
      <div class="fline">${L('🎲 Su locura:','🎲 Wildest bet:')} <b>${big}</b></div>`;
    grid.appendChild(card);
  });
  s.appendChild(grid);
  inp.addEventListener('input', () => { const q = inp.value.toLowerCase().trim();
    grid.querySelectorAll('.ficha').forEach(f => { f.style.display = f.dataset.name.includes(q) ? '' : 'none'; }); });
}

/* ---- PREMIOS ---- */
function buildPremios(){
  const s = section('premios', L('08 · Palmarés','08 · Awards'),
    L('El palmarés de la porra 🏅',"The pool's hall of fame 🏅"),
    L('Los títulos honoríficos de esta edición.',"This edition's honorary titles."));
  const a = D.awards;
  const items = [
    ['🐺', L('El Rebelde','The Maverick'), a.rebelde.name, L('índice '+a.rebelde.index+'/100','index '+a.rebelde.index+'/100')],
    ['🐑', L('El Borrego','The Sheep'), a.borrego.name, L('el más previsible ('+a.borrego.index+'/100)','most predictable ('+a.borrego.index+'/100)')],
    ['💞', L('Almas Gemelas','Soulmates'), a.gemelas.a+' & '+a.gemelas.b, L(a.gemelas.pct+'% idénticos',a.gemelas.pct+'% identical')],
    ['🧊', L('Polos Opuestos','Opposites'), a.opuestos.a+' & '+a.opuestos.b, L('distancia '+a.opuestos.dist,'distance '+a.opuestos.dist)],
    ['🥅', L('El Goleador','The Goal Machine'), a.goleador.name, L(a.goleador.avg+' goles/partido',a.goleador.avg+' goals/match')],
    ['🔒', L('El Cerrojo','The Parked Bus'), a.cerrojo.name, L(a.cerrojo.avg+' goles/partido',a.cerrojo.avg+' goals/match')],
    ['🤝', L('Rey del Empate','Draw King'), a.empate.name, L(a.empate.pct+'% en tablas',a.empate.pct+'% drawn')],
    ['🎯', L('Lobo Solitario','Lone Wolf'), a.lobo.name, L(a.lobo.count+' apuestas únicas',a.lobo.count+' unique bets')],
  ];
  const grid = el('div','grid g4 reveal');
  items.forEach(([em,ti,wn,dt]) => grid.appendChild(el('div','award',
    `<div class="em">${em}</div><div class="ti">${ti}</div><div class="wn">${esc(wn)}</div><div class="dt">${dt}</div>`)));
  s.appendChild(grid);
}

/* ---- ACIERTOS ---- */
function buildAciertos(){
  const s = section('aciertos', L('📈 Ranking','📈 Ranking'),
    L('Aciertos en directo ⚽','Live scoring ⚽'),
    L('Cuando empiece a rodar el balón y rellenes los resultados reales en el Excel, aquí saldrá el ranking de aciertos.',
      'Once the ball starts rolling and you fill in the real results in the Excel, the scoring ranking appears here.'));
  if(!D.live){
    s.appendChild(el('div','card teaser reveal',
      `<div class="em">⚽️</div><h3 style="margin:10px 0 6px">${L('El Mundial aún no ha empezado',"The World Cup hasn't kicked off yet")}</h3>
       <p class="muted">${L('Rellena los marcadores reales en la pestaña <b>Real results</b> del Excel y vuelve a generar el dashboard: aparecerá aquí el ranking de aciertos (signo +2, marcador exacto +4).','Fill in the real scores in the <b>Real results</b> tab of the Excel and regenerate the dashboard: the scoring ranking shows up here (correct outcome +2, exact score +4).')}</p>`));
    return;
  }
  const t = D.live.table, mx = t[0].pts || 1;
  const variant = currentRankingVariant();
  if(variant){
    buildRankingPrototype(s, variant);
    return;
  }
  const list = el('div','reveal');
  t.forEach((r,i) => list.appendChild(el('div','bar-row',
    `<div class="bar-rank">${i+1}</div><div class="bar-name">${i===0?'👑 ':''}${esc(r.name)} <span class="muted" style="font-size:.78rem">(${r.exact} ${L('plenos','exact')} · ${r.sign} ${L('signos','outcomes')})</span></div>
     <div class="bar-track"><div class="bar-fill gold" data-w="${(r.pts/mx*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${r.pts}">0</div>`)));
  s.appendChild(list);
}

/* ---- FOOTER ---- */
function buildFooter(){
  const f = el('footer');
  f.innerHTML = `<span class="brand">${logo}</span> ${L('Porra Mundial 2026 · análisis de '+N+' quinielas · generado el __DATE__.','World Cup 2026 Pool · analysis of '+N+' predictions · generated on __DATE__.')}
    <br><span style="opacity:.6">${L('Hecho con cariño para la oficina. Tipografía corporativa real: Garnett (de pago); aquí usamos Space Grotesk + Inter como sustitutas libres.','Made with love for the office. Real brand typeface: Garnett (paid); here we use Space Grotesk + Inter as free stand-ins.')}</span>`;
  wrap.appendChild(f);
}

/* ---- OBSERVERS ---- */
function observeAll(){
  const io = new IntersectionObserver((es,o) => { es.forEach(e => { if(e.isIntersecting){
    e.target.classList.add('in');
    e.target.querySelectorAll && e.target.querySelectorAll('[data-count]').forEach(animateCount);
    e.target.querySelectorAll && e.target.querySelectorAll('.bar-fill').forEach(b => { b.style.width = (b.dataset.w||0)+'%'; });
    o.unobserve(e.target);
  }}); }, {threshold:.12});
  wrap.querySelectorAll('.reveal').forEach(n => io.observe(n));
  const navlinks = {}; rail.querySelectorAll('a[data-id]').forEach(a => navlinks[a.dataset.id] = a);
  const navio = new IntersectionObserver(es => { es.forEach(e => { if(e.isIntersecting){
    Object.values(navlinks).forEach(a => a.classList.remove('active'));
    if(navlinks[e.target.id]) navlinks[e.target.id].classList.add('active');
  }}); }, {rootMargin:'-45% 0px -50% 0px'});
  wrap.querySelectorAll('header.hero,section.sec').forEach(sec => navio.observe(sec));
}

/* ---- BUILD / REBUILD ---- */
function rebuild(){
  clearRaceTimer();
  if(wrap) wrap.remove();
  wrap = el('div','wrap'); document.body.appendChild(wrap);
  renderRail();
  buildHoy(); buildAciertos(); buildUltimos(); buildHero(); buildRebeldia(); buildAfinidad(); buildEstilo(); buildFavoritos();
  buildPartidos(); buildLobo(); buildFichas(); buildPremios(); buildFooter();
  observeAll();
  renderPrototypeSwitcher();
}
rebuild();
"""

HTML_TEMPLATE = (
    '<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
    '<title>Porra Mundial 2026 · Reveni</title>\n'
    '<link rel="icon" href="favicon.svg" type="image/svg+xml">\n'
    '<link rel="preconnect" href="https://fonts.googleapis.com">\n'
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n'
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&'
    'family=Space+Grotesk:wght@500;700&display=swap" rel="stylesheet">\n'
    '<style>' + CSS + '</style>\n</head>\n<body>\n'
    '<div id="app"></div>\n'
    '<script>window.__PORRA__=__DATA__;</script>\n'
    '<script>' + JS + '</script>\n'
    '</body>\n</html>\n'
)


def load_logo():
    p = Path(LOGO_SVG_PATH)
    if not p.exists():
        return '<span class="wordmark">reveni</span>'
    svg = p.read_text(encoding="utf-8")
    return svg.replace('fill="#003C46"', 'fill="currentColor"')


def render_html(analytics, generated_on):
    data_json = json.dumps(analytics, ensure_ascii=False)
    logo = load_logo()
    return (
        HTML_TEMPLATE
        .replace("__DATA__", data_json)
        .replace("__LOGO__", logo)
        .replace("__DATE__", generated_on)
    )


if __name__ == "__main__":
    from datetime import date

    xlsx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    out = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUT
    data = parse_workbook(xlsx)
    analytics = compute(data)
    html = render_html(analytics, date.today().strftime("%d/%m/%Y"))
    Path(out).write_text(html, encoding="utf-8")
    print(f"OK · {data['n']} participantes · {len(data['matches'])} partidos · "
          f"resultados cargados: {len(data['results'])}")
    print(f"Dashboard escrito en: {out}  ({len(html) // 1024} KB)")
