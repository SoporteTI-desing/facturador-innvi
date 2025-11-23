// ========================= CONFIGURACIÓN EMISOR =========================
// Basado en tu CFDI de ejemplo de NOMAD GENETICS (RFC, régimen, CP, etc.)
const EMISOR = {
  nombre: "NOMAD GENETICS",
  rfc: "NIN190116954",
  regimenFiscal: "601",          // General de Ley Personas Morales
  direccion: "LA MORENA No. 811, Col. NARVARTE PONIENTE, CP 03020, CDMX, MÉXICO",
  cp: "03020"
};

// ========================= ESTADO GLOBAL =========================
let satCatalogs = null;

// empresa seleccionada
let empresaActual = "sanare"; // por defecto

// datos actuales en memoria (se usan en toda la app)
let clientes = [];
let productos = [];
let compras = [];
let conceptosFactura = [];

// claves de almacenamiento por empresa
const LS_KEYS = {
  clientesSanare: "factu_clientes_sanare",
  clientesNomad: "factu_clientes_nomad",
  productosSanare: "factu_productos_sanare",
  productosNomad: "factu_productos_nomad",
  comprasSanare: "factu_compras_sanare",
  comprasNomad: "factu_compras_nomad",
  empresaActual: "factu_empresa_actual",
  folioSanare: "factu_folio_sanare",
  folioNomad: "factu_folio_nomad"
};

// catálogos base cargados desde tus Excel
const CLIENTES_BASE_SANARE = [
  {
    "id": "sanare-cli-1",
    "nombre": "PUBLICO EN GENERAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "sanare-cli-2",
    "nombre": "Cliente Mostrador",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "sanare-cli-3",
    "nombre": "LUIS GABRIEL MONTES MEDINA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VERACRUZ",
    "email": ""
  },
  {
    "id": "sanare-cli-4",
    "nombre": "JOSE FERNANDO TORRES GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "XOCHITLA :MZ.3LT.44",
    "email": ""
  },
  {
    "id": "sanare-cli-5",
    "nombre": "VICTOR MANUEL ALLENDE TELLEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ROBLE",
    "email": ""
  },
  {
    "id": "sanare-cli-6",
    "nombre": "SANDRA PALACIOS MARQUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "RAFAEL ANGELPEÑA",
    "email": ""
  },
  {
    "id": "sanare-cli-7",
    "nombre": "MARCOS AUGUSTO PORTILLO CARRILLO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV 5 DE MAYO",
    "email": ""
  },
  {
    "id": "sanare-cli-8",
    "nombre": "CUITLAHUAC OMAR CASTRO VERDUGO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TERCERA OESTE",
    "email": ""
  },
  {
    "id": "sanare-cli-9",
    "nombre": "KAREN HERAS GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "DOCTOR AGUSTIN TORRES CRAVIOTO",
    "email": ""
  }
];
const CLIENTES_BASE_NOMAD = [
  {
    "id": "nomad-cli-1",
    "nombre": "PUBLICO GENERAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-2",
    "nombre": "SOHIN GENETICS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GUILLERMO GONZALEZ CAMARENA",
    "email": ""
  },
  {
    "id": "nomad-cli-3",
    "nombre": "SEGUROS INBURSA, S.A., GRUPO FINANCIERO INBURSA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA INSURGENTES SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-4",
    "nombre": "ALEJANDRO ZARCO MENDOZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PAOLO UCCELLO",
    "email": ""
  },
  {
    "id": "nomad-cli-5",
    "nombre": "AXA SEGUROS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV FELIX CUEVAS",
    "email": ""
  },
  {
    "id": "nomad-cli-6",
    "nombre": "SEGUROS SURA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA DE LOS INSURGENTES SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-7",
    "nombre": "PROMOTORA SN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CORAS",
    "email": ""
  },
  {
    "id": "nomad-cli-8",
    "nombre": "FRANCISCO JAVIER GONZALEZ BARRENA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LOMA DE PINAL DE AMOLES",
    "email": ""
  },
  {
    "id": "nomad-cli-9",
    "nombre": "OMAR YAIR CASTILLO CRUZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ANGELA PERALTA",
    "email": ""
  },
  {
    "id": "nomad-cli-10",
    "nombre": "MARIO ENRIQUE ESQUINCA MICELI",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "8A PONIENTE NORTE",
    "email": ""
  },
  {
    "id": "nomad-cli-11",
    "nombre": "GUILLERMINA SOFIA COPPEL KELLY",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ERNESTO COPPEL CAMPAÑA",
    "email": ""
  },
  {
    "id": "nomad-cli-12",
    "nombre": "GILBERTO CELIS LONGORIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TARASCO",
    "email": ""
  },
  {
    "id": "nomad-cli-13",
    "nombre": "CARLOS EDUARDO ZERON BENITEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE 5",
    "email": ""
  },
  {
    "id": "nomad-cli-14",
    "nombre": "MEDALFA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "EMILIANO ZAPATA BODEGA B2, B3, B4 PARQUE INDUSTRIAL LERMA",
    "email": ""
  },
  {
    "id": "nomad-cli-15",
    "nombre": "MARIANELA SANTOVEÑA RODRIGUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PROL. CALLE 7",
    "email": ""
  },
  {
    "id": "nomad-cli-16",
    "nombre": "MARIA ADRIANA DELABRA DEL RAZO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-17",
    "nombre": "CLINICA DIGITAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JESUS DEL MONTE",
    "email": ""
  },
  {
    "id": "nomad-cli-18",
    "nombre": "SYLVIA DIANA ORDOÑEZ Y AROESTE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "INSURGENTES SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-19",
    "nombre": "KARLO MAGNO DE JESUS GOMEZ GASCA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PLAYA CALETA",
    "email": ""
  },
  {
    "id": "nomad-cli-20",
    "nombre": "URSINA IBETT REYES GUTIERREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FORRAJES",
    "email": ""
  },
  {
    "id": "nomad-cli-21",
    "nombre": "GUILLERMO BRENIS TEJEDA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TEPEACA",
    "email": ""
  },
  {
    "id": "nomad-cli-22",
    "nombre": "JUAN CARLOS GARCIA CANO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CERRO DE LAS TORRES",
    "email": ""
  },
  {
    "id": "nomad-cli-23",
    "nombre": "CENTRO DE INFUSION PEDIATRICA INTEGRAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PERIFERICO SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-24",
    "nombre": "PUNTO PEN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VERSALLES",
    "email": ""
  },
  {
    "id": "nomad-cli-25",
    "nombre": "AGNI ATENEA TERRES RINCON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GRAL. BENJAMIN ARQUIMEDES",
    "email": ""
  },
  {
    "id": "nomad-cli-26",
    "nombre": "LOURDES ESTHER GUEVARA CASTAÑEDA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PITAGORAS 305-104",
    "email": ""
  },
  {
    "id": "nomad-cli-27",
    "nombre": "ABEL PEREZ LEDESMA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE HEROICO COLEGIO MILITAR S/N",
    "email": ""
  },
  {
    "id": "nomad-cli-28",
    "nombre": "JOSE RENE MANUEL ANGUIANO MARTINEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LIC. LUIS DONALDO COLOSIO MURRIETA",
    "email": ""
  },
  {
    "id": "nomad-cli-29",
    "nombre": "MARCELA NOEMI RAMOS GUILLEN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA ROBLES",
    "email": ""
  },
  {
    "id": "nomad-cli-30",
    "nombre": "GRUPO CUAUHTEMOC MOCTEZUMA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ALFONSO REYES NTE",
    "email": ""
  },
  {
    "id": "nomad-cli-31",
    "nombre": "FERNANDO UREÑA NUÑEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "NICOLAS BRAVO",
    "email": ""
  },
  {
    "id": "nomad-cli-32",
    "nombre": "MARIA ESTELA GONZALEZ SANCHEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA TAMAULIPAS",
    "email": ""
  },
  {
    "id": "nomad-cli-33",
    "nombre": "PUEBLICO EN GENERAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BEST DOCTOR INSURANCE 5201 BLUE LAGOON DRIVE SUITE 300 MIAMI  FL 33126",
    "email": ""
  },
  {
    "id": "nomad-cli-34",
    "nombre": "MYRIAM SERRA CASTELLANOS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CERRO DE LAS CAMPANAS 380",
    "email": ""
  },
  {
    "id": "nomad-cli-35",
    "nombre": "ADRIANA SALAZAR ANZURES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "80 MANZANA 24 LOTE 32",
    "email": ""
  },
  {
    "id": "nomad-cli-36",
    "nombre": "NUVIA ARELY NAVARRETE CERVANTES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CAMINO A LA VITELA KM 1",
    "email": ""
  },
  {
    "id": "nomad-cli-37",
    "nombre": "JESSICA GUERRERO ALVAREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CARRIL NTE A SN CRISTOBAL",
    "email": ""
  },
  {
    "id": "nomad-cli-38",
    "nombre": "RAUL CABALLERO LOPEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "RIO BALSAS 28",
    "email": ""
  },
  {
    "id": "nomad-cli-39",
    "nombre": "HECTOR MORALES ROJAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CAMINO AL DESIETO DE LOS LEONES 5232",
    "email": ""
  },
  {
    "id": "nomad-cli-40",
    "nombre": "PRODUCTOS ROCHE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CERRADA DE BEZARES",
    "email": ""
  },
  {
    "id": "nomad-cli-41",
    "nombre": "LESLEE JOVANA LOPEZ GUTIERREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GERARDO CARRASCO S/N",
    "email": ""
  },
  {
    "id": "nomad-cli-42",
    "nombre": "JAVIER TEJEDA VALLEJO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "EL ROSARIO",
    "email": ""
  },
  {
    "id": "nomad-cli-43",
    "nombre": "MARIA EUGENIA ESTRADA NOBLE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LAGUNA DE MAYRAN TORRE LONDRES DEPTO. 203",
    "email": ""
  },
  {
    "id": "nomad-cli-44",
    "nombre": "LUIS MANUEL ALVA VALDIVIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MAIZ",
    "email": ""
  },
  {
    "id": "nomad-cli-45",
    "nombre": "ANCORA INGENIERIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VIVEROS DE ASIS",
    "email": ""
  },
  {
    "id": "nomad-cli-46",
    "nombre": "PABLO CUARTO GONZALEZ RUBIO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "EJIDO SAN LORENZO TEZONCO",
    "email": ""
  },
  {
    "id": "nomad-cli-47",
    "nombre": "ENRIQUE ANTONIO SANTA CRUZ POLANCO Y BERUMEN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "21A",
    "email": ""
  },
  {
    "id": "nomad-cli-48",
    "nombre": "FERNANDO MATIAS MORENO YNTRIAGO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LLUVIA",
    "email": ""
  },
  {
    "id": "nomad-cli-49",
    "nombre": "ENRIQUE FERNANDO SALAZAR DIAZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PASEO DEL RIO",
    "email": ""
  },
  {
    "id": "nomad-cli-50",
    "nombre": "EDGAR ALBERTO SANCHEZ HERNANDEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VILLAS DE MERCURIO",
    "email": ""
  },
  {
    "id": "nomad-cli-51",
    "nombre": "CAMIORENTA MX",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CUAUHTEMOC",
    "email": ""
  },
  {
    "id": "nomad-cli-52",
    "nombre": "MIROSLAWA MAJA MISZALSKA HARASYMOWICZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "HERACLITO",
    "email": ""
  },
  {
    "id": "nomad-cli-53",
    "nombre": "JOSE ANTONIO DAVILA ORTIZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "HACIENDA DEL CIERVO",
    "email": ""
  },
  {
    "id": "nomad-cli-54",
    "nombre": "MARIA CELINDA GONZALEZ NAVARRETE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "POK TA POK",
    "email": ""
  },
  {
    "id": "nomad-cli-55",
    "nombre": "GUSTAVO DE LA SERNA CARDENAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PROLONGACION VASCO DE QUIROGA,  EDIFICIO E DEPTO. 1302",
    "email": ""
  },
  {
    "id": "nomad-cli-56",
    "nombre": "MIGUEL ANGEL HERRERA TORRES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PROGRESO",
    "email": ""
  },
  {
    "id": "nomad-cli-57",
    "nombre": "FERNANDO GARCIA GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ASTRON",
    "email": ""
  },
  {
    "id": "nomad-cli-58",
    "nombre": "JESUS ANTONIO AVILA GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ACTOPAN COLONIA  SANTA JULIA INFONAVIT",
    "email": ""
  },
  {
    "id": "nomad-cli-59",
    "nombre": "ESTEBAN BERNAL QUINTANA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PROLONGACION VICENTE GUERRERO",
    "email": ""
  },
  {
    "id": "nomad-cli-60",
    "nombre": "MARIA DEL CARMEN ARIAS POLLEDO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV DE LS TORRES",
    "email": ""
  },
  {
    "id": "nomad-cli-61",
    "nombre": "EDGAR ULISES ESPINOSA MORENO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE RIO LAJAS",
    "email": ""
  },
  {
    "id": "nomad-cli-62",
    "nombre": "ALVARO URQUIZA ESTRADA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GRANJA EL PILAR",
    "email": ""
  },
  {
    "id": "nomad-cli-63",
    "nombre": "JOSE ANTONIO BLANCO GAMBOA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "COLONIA MACUZARI",
    "email": ""
  },
  {
    "id": "nomad-cli-64",
    "nombre": "BIBIANA CAROLINA SANDOVAL ENRIQUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LUIS REYES SPINDOLA",
    "email": ""
  },
  {
    "id": "nomad-cli-65",
    "nombre": "FRANCISCO JOSE MICHAUS MONTES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "EJE CENTRAL LAZARO CARDENAS",
    "email": ""
  },
  {
    "id": "nomad-cli-66",
    "nombre": "MARIA ESQUIVEL  QUINTERO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ENRIQUE REBSAMEN",
    "email": ""
  },
  {
    "id": "nomad-cli-67",
    "nombre": "ANTONIO POMPA GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CANGREJO",
    "email": ""
  },
  {
    "id": "nomad-cli-68",
    "nombre": "URGENT 24",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TAMAULIPAS",
    "email": ""
  },
  {
    "id": "nomad-cli-69",
    "nombre": "MARIA DE LOURDES CRUZ RODRIGUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CIRCUNVALACION",
    "email": ""
  },
  {
    "id": "nomad-cli-70",
    "nombre": "FRANCISCO JAVIER GONGORA VAZQUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "55",
    "email": ""
  },
  {
    "id": "nomad-cli-71",
    "nombre": "CASA DE LA AMISTAD PARA NIÑOS CON CANCER",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ALDAMA",
    "email": ""
  },
  {
    "id": "nomad-cli-72",
    "nombre": "ROBERTO ESCOBAR ARIAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "Esperanza",
    "email": ""
  },
  {
    "id": "nomad-cli-73",
    "nombre": "UNIVERSIDAD NACIONAL AUTONOMA DE MEXICO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-74",
    "nombre": "CARLOS ALBERTO SILICEO CURIEL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ANTONIO DE MENDOZA",
    "email": ""
  },
  {
    "id": "nomad-cli-75",
    "nombre": "PAULINA MONTSERRAT MOLINA TELLEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FELIX CUEVAS",
    "email": ""
  },
  {
    "id": "nomad-cli-76",
    "nombre": "JULIETA GUTIERREZ ZAMORA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SALINEROS",
    "email": ""
  },
  {
    "id": "nomad-cli-77",
    "nombre": "SANIDAD DIGITAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOSE RUBEN ROMERO",
    "email": ""
  },
  {
    "id": "nomad-cli-78",
    "nombre": "PEDRO IVAN CORREA BENAVENTE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GUADALUPE",
    "email": ""
  },
  {
    "id": "nomad-cli-79",
    "nombre": "ROGERIO ARMANDO CHACON SOLIS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "QUINTANA ROO  PISO 5 CONSULTORIO 507",
    "email": ""
  },
  {
    "id": "nomad-cli-80",
    "nombre": "ROMAN FERNANDO SOLARES TERAN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ANSELMO DE LA PORTILLA UNIDAD 4 EDIFICIO 6 DEPARTAMENTO 63",
    "email": ""
  },
  {
    "id": "nomad-cli-81",
    "nombre": "SERVICIOS ESTATALES DE SALUD",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CHAPULTEPEC",
    "email": ""
  },
  {
    "id": "nomad-cli-82",
    "nombre": "VALERIA GUADALUPE MILLAN ENCARNACION",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOAQUIN VELAZQUEZ DE LEON",
    "email": ""
  },
  {
    "id": "nomad-cli-83",
    "nombre": "NUCLEO DE ATENCION MATERNA Y SALUD FETAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MONTECITO",
    "email": ""
  },
  {
    "id": "nomad-cli-84",
    "nombre": "TERE ALMAZAN MASON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ESTRELLA EDIFICIO A INTERIOR 111",
    "email": ""
  },
  {
    "id": "nomad-cli-85",
    "nombre": "ABAVEN CONSULTING",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SAN BENITO",
    "email": ""
  },
  {
    "id": "nomad-cli-86",
    "nombre": "EDGAR ALEJANDRO CRUZ GUZMAN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-87",
    "nombre": "ROGELIO MARTINEZ GOMEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VICTORIA",
    "email": ""
  },
  {
    "id": "nomad-cli-88",
    "nombre": "ESTUDIO BOSCO SODI",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-89",
    "nombre": "ALEJANDRA IBARRA CHAOUL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-90",
    "nombre": "PERLA YOLANDA FABRO TREJO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-91",
    "nombre": "CENTRO DE ATENCION MATERNO FETAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PERIFERICO SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-92",
    "nombre": "MASSIEL MARLEN BETANZOS NUÑEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-93",
    "nombre": "LORENZO JOSE VELANDIA MONTES DE OCA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "NICOLAS SAN JUAN",
    "email": ""
  },
  {
    "id": "nomad-cli-94",
    "nombre": "CINTHYA OCHOA VEGA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "NICOLAU INTERIOR",
    "email": ""
  },
  {
    "id": "nomad-cli-95",
    "nombre": "CYMARKETING",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LAGUNA DEL CARMEN",
    "email": ""
  },
  {
    "id": "nomad-cli-96",
    "nombre": "RAUL ARMANDO SOBRINO DE LA CERDA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PROVIDENCIA",
    "email": ""
  },
  {
    "id": "nomad-cli-97",
    "nombre": "ARTURO FREG TENORIO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MONTECILLO",
    "email": ""
  },
  {
    "id": "nomad-cli-98",
    "nombre": "EDITH ELIZABETH MORA HERNANDEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV DR JOSE MARIA VERTIZ",
    "email": ""
  },
  {
    "id": "nomad-cli-99",
    "nombre": "DIANA PERAFAN VILLEGAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CENTRAL  EDIFICIO 2 CASA 1",
    "email": ""
  },
  {
    "id": "nomad-cli-100",
    "nombre": "LAURA VALERIA RAMIREZ GILES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TECAX MANZANA 45 LOTE 09",
    "email": ""
  },
  {
    "id": "nomad-cli-101",
    "nombre": "MONTSERRAT GUZMAN GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-102",
    "nombre": "JACQUELINE JHOANNA MAYA QUIÑONES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GALLOS 4",
    "email": ""
  },
  {
    "id": "nomad-cli-103",
    "nombre": "CINTHYA GABRIELA GUTIERREZ ZUÑIGA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "HACIENDA DE LA LLAVE",
    "email": ""
  },
  {
    "id": "nomad-cli-104",
    "nombre": "KARLA ELIZABETH OLVERA PARADA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BAZOFERTEMANZANA3",
    "email": ""
  },
  {
    "id": "nomad-cli-105",
    "nombre": "UNIDAD AVANZADA DE DIAGNOSTICO MATERNO FETAL CUERNAVACA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CHILPANCINGO",
    "email": ""
  },
  {
    "id": "nomad-cli-106",
    "nombre": "MARIA FERNANDA GARCIA TRABADO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "XICOTENCATL",
    "email": ""
  },
  {
    "id": "nomad-cli-107",
    "nombre": "EVELYN MORENO MEDINA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BOSQUE DE NIAMEY",
    "email": ""
  },
  {
    "id": "nomad-cli-108",
    "nombre": "FRANCISCO JAVIER MORENO MORALES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOSE MARIA MORELOS Y PAVON",
    "email": ""
  },
  {
    "id": "nomad-cli-109",
    "nombre": "CARLOS UTRERA MENDEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GENERAL INGNACIO DE LA LLAVE NTE",
    "email": ""
  },
  {
    "id": "nomad-cli-110",
    "nombre": "JAQUELINE GARCIA GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "4A AVENIDA",
    "email": ""
  },
  {
    "id": "nomad-cli-111",
    "nombre": "IVETTE GABRIELA LOZANO FLORES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "NORTE 24 A",
    "email": ""
  },
  {
    "id": "nomad-cli-112",
    "nombre": "CATALINA FERNANDA PADRES ROJO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVE. SANTA FE 425 EDIFICIO A",
    "email": ""
  },
  {
    "id": "nomad-cli-113",
    "nombre": "JESUS ARTURO SALAS SALAZAR",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MANZANO",
    "email": ""
  },
  {
    "id": "nomad-cli-114",
    "nombre": "INMOBILIARIA ORANGE COUNTRY",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "REFORMA SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-115",
    "nombre": "INGENIERIA INTEGRAL DE VALUACION",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ACORDADA",
    "email": ""
  },
  {
    "id": "nomad-cli-116",
    "nombre": "ANGELA LEZAMA CASTRO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "IGNACIO ALLENDE",
    "email": ""
  },
  {
    "id": "nomad-cli-117",
    "nombre": "ORNELLA AGOSTINA SAURO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "OBREROMUNDIAL",
    "email": ""
  },
  {
    "id": "nomad-cli-118",
    "nombre": "ULTRASCANMX",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "XOLA",
    "email": ""
  },
  {
    "id": "nomad-cli-119",
    "nombre": "ELVIA MORENO SANCHEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LAGUNA DE TERMINOS",
    "email": ""
  },
  {
    "id": "nomad-cli-120",
    "nombre": "ANA PAULA RODRIGUEZ GALVEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LA JOYA 62",
    "email": ""
  },
  {
    "id": "nomad-cli-121",
    "nombre": "CARLOS JUAN RABINDRANATH CANTU HERRERA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOSE MARIA MORELOS",
    "email": ""
  },
  {
    "id": "nomad-cli-122",
    "nombre": "ARACELI LEON SANCHEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA DE LA LUNA",
    "email": ""
  },
  {
    "id": "nomad-cli-123",
    "nombre": "ALDO HERNAN JARAMILLO ROMERO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE ANDADOR CEIBA",
    "email": ""
  },
  {
    "id": "nomad-cli-124",
    "nombre": "OSWALDO ALEJANDRO LOPEZ ARELLANOS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CANDELARIA",
    "email": ""
  },
  {
    "id": "nomad-cli-125",
    "nombre": "MARLEN JOCELYN LUNA RIVERA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LAURELES",
    "email": ""
  },
  {
    "id": "nomad-cli-126",
    "nombre": "FRIDA HARARI CHEREM",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FUENTES DE MOLINOS TORRE A INT 102",
    "email": ""
  },
  {
    "id": "nomad-cli-127",
    "nombre": "YESENIA ELIZABETH HUIZAR VILLALOBOS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ANDADOR LLUVIA",
    "email": ""
  },
  {
    "id": "nomad-cli-128",
    "nombre": "AVID SAUL PACHECO GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "RAMON TORRES",
    "email": ""
  },
  {
    "id": "nomad-cli-129",
    "nombre": "CLAUDIA ANDREA SANCHEZ RODRIGUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LE EROS",
    "email": ""
  },
  {
    "id": "nomad-cli-130",
    "nombre": "DANIELA ARIAS ARANDA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MONACO",
    "email": ""
  },
  {
    "id": "nomad-cli-131",
    "nombre": "INTEGRA SALUD Y SERVICIOS MEDICOS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MIGUEL LERDO DE TEJADA",
    "email": ""
  },
  {
    "id": "nomad-cli-132",
    "nombre": "ESTEFANIA ZAVALA LOPEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PASEO VIZCARRA",
    "email": ""
  },
  {
    "id": "nomad-cli-133",
    "nombre": "MARY CARMEN CUEVAS REYES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SUR 143",
    "email": ""
  },
  {
    "id": "nomad-cli-134",
    "nombre": "ALBERTO BELTRAN CAMPOS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ORQUIDEAS",
    "email": ""
  },
  {
    "id": "nomad-cli-135",
    "nombre": "MARIA EUGENIA GONZALEZ VARGAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PORTAL MORELOS",
    "email": ""
  },
  {
    "id": "nomad-cli-136",
    "nombre": "EDGAR ABEL MARQUEZ GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "OLIVOS",
    "email": ""
  },
  {
    "id": "nomad-cli-137",
    "nombre": "ANA TERESA MEDEL GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "95 PTE ALTOS",
    "email": ""
  },
  {
    "id": "nomad-cli-138",
    "nombre": "LUIS FERNANDO GAMA MUÑOZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AUXILIADORAS",
    "email": ""
  },
  {
    "id": "nomad-cli-139",
    "nombre": "ALEJANDRO MEZA DE ALBA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-140",
    "nombre": "MARIA BLANCA DIAZ ALVAREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MARGARITAS",
    "email": ""
  },
  {
    "id": "nomad-cli-141",
    "nombre": "PAN - AMERICAN MEXICO, COMPAÑIA DE SEGUROS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV. PASEO DE LA REFORMA 411 SUITE1501",
    "email": ""
  },
  {
    "id": "nomad-cli-142",
    "nombre": "EFREN OCTAVIANO MENDEZ TREVILLA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CAMINO A SANTA TERESA",
    "email": ""
  },
  {
    "id": "nomad-cli-143",
    "nombre": "ALONDRA SELENE ROMERO GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MEXICO Y ABASOLO",
    "email": ""
  },
  {
    "id": "nomad-cli-144",
    "nombre": "YOLANDA MARIA GUADALUPE CORTES ORDOÑEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PESTALOZZI",
    "email": ""
  },
  {
    "id": "nomad-cli-145",
    "nombre": "KEVIN DANIEL GARCIA OLGUIN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VALLE DE TULA MANZANA 11 LOTE 1",
    "email": ""
  },
  {
    "id": "nomad-cli-146",
    "nombre": "LUIS ANTONIO FUENTES SOBREYRA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PLANTA TINGAMBATO",
    "email": ""
  },
  {
    "id": "nomad-cli-147",
    "nombre": "PAULO CESAR ESPITIA JAIME",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ANTONIO DE MENDOZA",
    "email": ""
  },
  {
    "id": "nomad-cli-148",
    "nombre": "ELIA IXEL APODACA CHAVEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ARENAL",
    "email": ""
  },
  {
    "id": "nomad-cli-149",
    "nombre": "VICENTE VARA CELIS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SAN VICENTE",
    "email": ""
  },
  {
    "id": "nomad-cli-150",
    "nombre": "PAULETTE FRAGOSO SABBAGH",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOJUTLA",
    "email": ""
  },
  {
    "id": "nomad-cli-151",
    "nombre": "NAOMI SANTANA JIMENEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "POPOCATEPETL",
    "email": ""
  },
  {
    "id": "nomad-cli-152",
    "nombre": "JUAN CARLOS BUSTAMANTE OGANDO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "RIO MANILA",
    "email": ""
  },
  {
    "id": "nomad-cli-153",
    "nombre": "OSCAR OMAR HERNANDEZ MARTINEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FLOR DE NUBE MZ. 24 LT.3",
    "email": ""
  },
  {
    "id": "nomad-cli-154",
    "nombre": "MANUEL ABRAHAM VALDERRAMA YAPOR",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-155",
    "nombre": "SILVIA BERNARDETE REYNOSO ESPARZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VERACRUZ",
    "email": ""
  },
  {
    "id": "nomad-cli-156",
    "nombre": "EVA PATRICIA AVILES GOMEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "NOGAL",
    "email": ""
  },
  {
    "id": "nomad-cli-157",
    "nombre": "JONATHAN ISRAEL FONSECA MENDEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PUERTE CAMPECHE",
    "email": ""
  },
  {
    "id": "nomad-cli-158",
    "nombre": "RAUL ALBERTO RUIZ DIAZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TUXPAN TAMIAHUA",
    "email": ""
  },
  {
    "id": "nomad-cli-159",
    "nombre": "JUAN OMAR AVILES CORONA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "DIAGONAL 1",
    "email": ""
  },
  {
    "id": "nomad-cli-160",
    "nombre": "JOSE SANCHEZ SALGADO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ROTOGRABADOS",
    "email": ""
  },
  {
    "id": "nomad-cli-161",
    "nombre": "SILVIA IBETH GONZALEZ LOPEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA CONSTITUYENTES 1076 EDIFICIO PINOS",
    "email": ""
  },
  {
    "id": "nomad-cli-162",
    "nombre": "KEYDY AIRY ORDAZ GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PASEO DEL HALCON",
    "email": ""
  },
  {
    "id": "nomad-cli-163",
    "nombre": "CARLOTTA ELISABETH OETLING VON REDEN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SONORA",
    "email": ""
  },
  {
    "id": "nomad-cli-164",
    "nombre": "BLANCA DEL ROCIO MORENO AVILA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ESTADO DE MEXICO",
    "email": ""
  },
  {
    "id": "nomad-cli-165",
    "nombre": "ALBERTO EFRAIN CERVANTES ESPINOSA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JUAREZ",
    "email": ""
  },
  {
    "id": "nomad-cli-166",
    "nombre": "NANCY CANSECO RODAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TLAHUAC",
    "email": ""
  },
  {
    "id": "nomad-cli-167",
    "nombre": "CYNTHIA RODRIGUEZ DE JESUS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE MIGUEL BARRAGAN MANZANA 65 LOTE 7",
    "email": ""
  },
  {
    "id": "nomad-cli-168",
    "nombre": "LORENA MARTINEZ GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CAMINO A SANTA TERESA",
    "email": ""
  },
  {
    "id": "nomad-cli-169",
    "nombre": "HORACIO ZALDIVAR GAMBOA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "COOK",
    "email": ""
  },
  {
    "id": "nomad-cli-170",
    "nombre": "VICTOR MANUEL TELLEZ MEDINA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "DR. ANTONIO HERNANDEZ",
    "email": ""
  },
  {
    "id": "nomad-cli-171",
    "nombre": "GALDINA GARCIA LUDLOW",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "INSURGENTES",
    "email": ""
  },
  {
    "id": "nomad-cli-172",
    "nombre": "HEALTWAY",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "HERMANOS SERDAN",
    "email": ""
  },
  {
    "id": "nomad-cli-173",
    "nombre": "FERNANDO MONTOYA VELA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LEONARDO DE VINCI",
    "email": ""
  },
  {
    "id": "nomad-cli-174",
    "nombre": "ELIZABETH BARRIENTOS SANCHEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PRIVADA DEL VALLE",
    "email": ""
  },
  {
    "id": "nomad-cli-175",
    "nombre": "VIRGINIA PAOLA ROY LUZARRAGA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ARENAL",
    "email": ""
  },
  {
    "id": "nomad-cli-176",
    "nombre": "CHARLY LESGOURGUES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOSE MARTI",
    "email": ""
  },
  {
    "id": "nomad-cli-177",
    "nombre": "OSCAR ALEJANDRO SANCHEZ PEREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PRIVADA SOLES",
    "email": ""
  },
  {
    "id": "nomad-cli-178",
    "nombre": "XOCHITL YASMIN SANCHEZ GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE 11",
    "email": ""
  },
  {
    "id": "nomad-cli-179",
    "nombre": "FLORENCIO GARCIA ZARAIN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "11 SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-180",
    "nombre": "RODRIGO BARROSO CORONA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "OJO DE AGUA",
    "email": ""
  },
  {
    "id": "nomad-cli-181",
    "nombre": "DANIELA PLAZA HERNANDEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ORIENTE106",
    "email": ""
  },
  {
    "id": "nomad-cli-182",
    "nombre": "LUIS MARCOS CISNEROS FUENTES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "C 14 NORTE #1005, BARRIO DE SAN JOSE",
    "email": ""
  },
  {
    "id": "nomad-cli-183",
    "nombre": "MARCO ANTONIO GURROLA GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ARRAYAN",
    "email": ""
  },
  {
    "id": "nomad-cli-184",
    "nombre": "EDGAR STAMER GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "HACIENDA DE LAS PALMAS",
    "email": ""
  },
  {
    "id": "nomad-cli-185",
    "nombre": "UNIDAD ONCOLOGICA AMBULATORIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GONZALEZ ORTEGA",
    "email": ""
  },
  {
    "id": "nomad-cli-186",
    "nombre": "JOSE ANTONIO ORTIZ VALENCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ORIENTE 10",
    "email": ""
  },
  {
    "id": "nomad-cli-187",
    "nombre": "ELIZABETH BARANDA DELGADO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "RECREO",
    "email": ""
  },
  {
    "id": "nomad-cli-188",
    "nombre": "AVELINO GONZALEZ GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA HIDALGO",
    "email": ""
  },
  {
    "id": "nomad-cli-189",
    "nombre": "ALEJANDRO RIEFKOHL CRAULES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MONTEEVEREST",
    "email": ""
  },
  {
    "id": "nomad-cli-190",
    "nombre": "ROSALINDA FRANCO SOTO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FRANCIA",
    "email": ""
  },
  {
    "id": "nomad-cli-191",
    "nombre": "MARIA AURORA VIDAL ESTRADA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE 16",
    "email": ""
  },
  {
    "id": "nomad-cli-192",
    "nombre": "JOSE ALBERTO MIRANDA CASILLAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MORELOS",
    "email": ""
  },
  {
    "id": "nomad-cli-193",
    "nombre": "ERNESTINA SALINAS GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FRAGIL VISTA",
    "email": ""
  },
  {
    "id": "nomad-cli-194",
    "nombre": "IVONNE MORENO CALDERON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA ARAUCARIAS",
    "email": ""
  },
  {
    "id": "nomad-cli-195",
    "nombre": "VICTOR OSCAR VILLAFAÑE FERRER",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "IBSEN",
    "email": ""
  },
  {
    "id": "nomad-cli-196",
    "nombre": "CARMEN HERNANDEZ ESPINOSA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CIPRES",
    "email": ""
  },
  {
    "id": "nomad-cli-197",
    "nombre": "PAOLA STEPHANIE GONZALEZ LOZANO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BENJAMIN FRANKLIN",
    "email": ""
  },
  {
    "id": "nomad-cli-198",
    "nombre": "SONIA ELIZABETH MALDONADO CABALLERO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JULIO JIMENEZ RUEDA",
    "email": ""
  },
  {
    "id": "nomad-cli-199",
    "nombre": "MARIA ELENA MORENO DIAZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LA GUADALUPANA",
    "email": ""
  },
  {
    "id": "nomad-cli-200",
    "nombre": "MARIA EDITH MORALES ESCOBAR",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PROLONGACION ALAMOS, EDIFICIO G",
    "email": ""
  },
  {
    "id": "nomad-cli-201",
    "nombre": "PREVEM SEGUROS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "",
    "email": ""
  },
  {
    "id": "nomad-cli-202",
    "nombre": "RICARDO PAREDES LUCIO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PUERTO GUAYMAS 43",
    "email": ""
  },
  {
    "id": "nomad-cli-203",
    "nombre": "MARIA LUISA NARDONE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "HACIENDA EL ROCIO",
    "email": ""
  },
  {
    "id": "nomad-cli-204",
    "nombre": "JUAN RODOLFO SANCHEZ GOMEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SITIO DE CUAUTLA",
    "email": ""
  },
  {
    "id": "nomad-cli-205",
    "nombre": "FLOR MARIA ESCAMILLA MALAGON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FUERTE DE LORETO",
    "email": ""
  },
  {
    "id": "nomad-cli-206",
    "nombre": "MOISES RAYEK KHABIE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SIERRA NEVADA",
    "email": ""
  },
  {
    "id": "nomad-cli-207",
    "nombre": "BEST DOCTORS INSURANCE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "5201 BLUE LAGOON DRIVE SITE 300 MIAMI, FL 33126",
    "email": ""
  },
  {
    "id": "nomad-cli-208",
    "nombre": "GERARDO CASTORENA ROJI",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ALICA",
    "email": ""
  },
  {
    "id": "nomad-cli-209",
    "nombre": "MARTIN GARCIA MARQUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CARRETERA TIJUANA TECATE",
    "email": ""
  },
  {
    "id": "nomad-cli-210",
    "nombre": "ALFONSO BARBA GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "DR. DIAZ DE  LEON",
    "email": ""
  },
  {
    "id": "nomad-cli-211",
    "nombre": "MARIA DEL ROSARIO SANDOVAL TINOCO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PRIVADA YUCATAN",
    "email": ""
  },
  {
    "id": "nomad-cli-212",
    "nombre": "ANGEL GERARDO DIAZ ACOSTA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GUILLERMOGONZALEZCAMARENA",
    "email": ""
  },
  {
    "id": "nomad-cli-213",
    "nombre": "JORGE ALBERTO SOBREVILLA GARZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE ALAMO",
    "email": ""
  },
  {
    "id": "nomad-cli-214",
    "nombre": "NORMA AVILA RAMIREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PROLONGACIONZARAGOZA",
    "email": ""
  },
  {
    "id": "nomad-cli-215",
    "nombre": "ALFONSO SALOMON MASRI DIWAN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "16 DE SEPTIEMBRE, BODEGA 5 PISO1",
    "email": ""
  },
  {
    "id": "nomad-cli-216",
    "nombre": "MARTHA CECILIA MARCOS ASSAD",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOSE VANSCONCELOS",
    "email": ""
  },
  {
    "id": "nomad-cli-217",
    "nombre": "MARIA FERNANDA OSORIO CARRILLO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "NAUHOLIN",
    "email": ""
  },
  {
    "id": "nomad-cli-218",
    "nombre": "ANA LUISA GUTIERREZ GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "21 PONIENTE DEPARTAMENTO 201",
    "email": ""
  },
  {
    "id": "nomad-cli-219",
    "nombre": "ASTRID NUÑEZ SALGADO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "INDEPENDENCIA",
    "email": ""
  },
  {
    "id": "nomad-cli-220",
    "nombre": "YAD LA JOLEH",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PRESIDENTE MASARIK",
    "email": ""
  },
  {
    "id": "nomad-cli-221",
    "nombre": "NORA PALOMA AMEZOLA VILLALPANDO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ANDEN 6, BODEGA 14",
    "email": ""
  },
  {
    "id": "nomad-cli-222",
    "nombre": "FRANCISCO RAUL VARGAS PINZON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BLVDMAVILAC",
    "email": ""
  },
  {
    "id": "nomad-cli-223",
    "nombre": "MARIA ASUNCION RAMIREZ FRIAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE CUATRO",
    "email": ""
  },
  {
    "id": "nomad-cli-224",
    "nombre": "TERESITA DEL NIÑO JESUS CORTES LLAMOSA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BENITO JUAREZ",
    "email": ""
  },
  {
    "id": "nomad-cli-225",
    "nombre": "MARIA HINESTROSA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SUDERMANN",
    "email": ""
  },
  {
    "id": "nomad-cli-226",
    "nombre": "FLORINA ENRIQUEZ LOPEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "16 SUR ORIENTE",
    "email": ""
  },
  {
    "id": "nomad-cli-227",
    "nombre": "MARIA DEL CARMEN ABASCAL ALVAREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FUENTEDEQUIJOTE",
    "email": ""
  },
  {
    "id": "nomad-cli-228",
    "nombre": "CLAUDIA ESTHELA MARTINEZ DE LA GARZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PALESTINA",
    "email": ""
  },
  {
    "id": "nomad-cli-229",
    "nombre": "JOSE FRANCISCO ALEXANDER MEZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA MOCTEZUMA",
    "email": ""
  },
  {
    "id": "nomad-cli-230",
    "nombre": "FRANK WOLBERG RODRIGUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE SAMARIA",
    "email": ""
  },
  {
    "id": "nomad-cli-231",
    "nombre": "ENRIQUE PEREZ CHAVEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CUAUTLA",
    "email": ""
  },
  {
    "id": "nomad-cli-232",
    "nombre": "JUAN MIGUEL ZATARAIN TISNADO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDAPASEODELOSAILES",
    "email": ""
  },
  {
    "id": "nomad-cli-233",
    "nombre": "SUSANA COUTTOLENC GARCIA JURADO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VERGEL",
    "email": ""
  },
  {
    "id": "nomad-cli-234",
    "nombre": "BEATRIZ CORDERO RODRIGUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "307",
    "email": ""
  },
  {
    "id": "nomad-cli-235",
    "nombre": "J. REMEDIOS MAJIN RODRIGUEZ ESCALANTE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MORELOS",
    "email": ""
  },
  {
    "id": "nomad-cli-236",
    "nombre": "PEDRO MORA BERMEJO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA 5 DE MAYO",
    "email": ""
  },
  {
    "id": "nomad-cli-237",
    "nombre": "BEATRIZ RIVERA VAZQUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "OTE 168",
    "email": ""
  },
  {
    "id": "nomad-cli-238",
    "nombre": "ELIO ESPINOSA RODRIGUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE AMORES",
    "email": ""
  },
  {
    "id": "nomad-cli-239",
    "nombre": "GRUPO NACIONAL PROVINCIAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA CERRO DE LAS TORRES",
    "email": ""
  },
  {
    "id": "nomad-cli-240",
    "nombre": "DIEGO GOMEZ HARO KATZNELSON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MAZATLAN",
    "email": ""
  },
  {
    "id": "nomad-cli-241",
    "nombre": "ANA HILDA SAMANO GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA DE LOS PELICANOS",
    "email": ""
  },
  {
    "id": "nomad-cli-242",
    "nombre": "AVRAHAM ASHKENAZI TAVASHI",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV. DE LOS BOSQUES",
    "email": ""
  },
  {
    "id": "nomad-cli-243",
    "nombre": "CLAUDIA YAZMIN LOPEZ MONZON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JENUFA",
    "email": ""
  },
  {
    "id": "nomad-cli-244",
    "nombre": "MEXICO INTEGRATIVE ONCOLOGY PROJECT",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ALICA",
    "email": ""
  },
  {
    "id": "nomad-cli-245",
    "nombre": "MARIA ANTONIA ANICETO PADILLA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "NORIA",
    "email": ""
  },
  {
    "id": "nomad-cli-246",
    "nombre": "MARIA ELENA CANAAN GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOSE MARIA MACHUCA",
    "email": ""
  },
  {
    "id": "nomad-cli-247",
    "nombre": "SERGIO ANTONIO BAYARDO TOLEDO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE SAN ERNESTO",
    "email": ""
  },
  {
    "id": "nomad-cli-248",
    "nombre": "ADOLFO ROMAN JAUREGUY RODRIGUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TACNA",
    "email": ""
  },
  {
    "id": "nomad-cli-249",
    "nombre": "JOSE LUIS MENCHELLI JIMENEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ALVARO OBREGON",
    "email": ""
  },
  {
    "id": "nomad-cli-250",
    "nombre": "DAVID AGAMI HAIAT",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV. INGENIEROS MILITARES",
    "email": ""
  },
  {
    "id": "nomad-cli-251",
    "nombre": "MARTHA REYNOSO ROBLES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AGUILA",
    "email": ""
  },
  {
    "id": "nomad-cli-252",
    "nombre": "ADALBERTO THOMAE LOPEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PASEO DE LOS FRESNOS",
    "email": ""
  },
  {
    "id": "nomad-cli-253",
    "nombre": "ALFREDO FEDERICO MORENO BENAVIDES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JOSE LOPEZ PORTILLO",
    "email": ""
  },
  {
    "id": "nomad-cli-254",
    "nombre": "CENTRO DE ONCOLOGIA DE PRECISION",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LA MORENA",
    "email": ""
  },
  {
    "id": "nomad-cli-255",
    "nombre": "BEATRIZ VEGA RUBIO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "Encino",
    "email": ""
  },
  {
    "id": "nomad-cli-256",
    "nombre": "LILIAN FRANCISCA ELIZABETH FRANK DIAZ DE SANDI",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "EFRAIN GONZALEZ LUNA",
    "email": ""
  },
  {
    "id": "nomad-cli-257",
    "nombre": "ALBERTO ESCOBELL RIOS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GALATONE",
    "email": ""
  },
  {
    "id": "nomad-cli-258",
    "nombre": "ALEJANDRO CASTIL VAZQUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CONSTITUCION",
    "email": ""
  },
  {
    "id": "nomad-cli-259",
    "nombre": "PLAN SEGURO SA DE CV COMPAÑIA DE SEGUROS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "ADOLFO RUIZ CORTINEZ",
    "email": ""
  },
  {
    "id": "nomad-cli-260",
    "nombre": "MA. TERESA AGUSTINA VELASCO GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SAN MATEO",
    "email": ""
  },
  {
    "id": "nomad-cli-261",
    "nombre": "EDUARDO ALEJANDRO GONGORA RAMIREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PISA",
    "email": ""
  },
  {
    "id": "nomad-cli-262",
    "nombre": "FRANCISCO JAVIER GOMEZ PLASCENCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FRANCISCO VILLA",
    "email": ""
  },
  {
    "id": "nomad-cli-263",
    "nombre": "JESUS ANTONIO MARIN ROJAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PLUTARCO ELIAS CALLES",
    "email": ""
  },
  {
    "id": "nomad-cli-264",
    "nombre": "FERNANDO VAZQUEZ MAYA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PRESIDENTES",
    "email": ""
  },
  {
    "id": "nomad-cli-265",
    "nombre": "MARIA FERNANDA MAGALLANES QUIROZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PASEO ROSAS",
    "email": ""
  },
  {
    "id": "nomad-cli-266",
    "nombre": "JOSE GUILLERMO CAMPOS LIMON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PRIVADA 3 A SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-267",
    "nombre": "EUGENIO MANILLA CALDERON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "EJIDO LOS REYES",
    "email": ""
  },
  {
    "id": "nomad-cli-268",
    "nombre": "PRECISION DIAGNOSTICA DALICE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LUIS DONALDO COLOSIO",
    "email": ""
  },
  {
    "id": "nomad-cli-269",
    "nombre": "MARCELA ALINA FUENTES MILAN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BOSQUE REAL",
    "email": ""
  },
  {
    "id": "nomad-cli-270",
    "nombre": "EMILIA RIVERO RIVERA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CARRETERA MEXICO TEXCOCO",
    "email": ""
  },
  {
    "id": "nomad-cli-271",
    "nombre": "JOSE ANTONIO JIMENEZ CORTEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PUEBLA",
    "email": ""
  },
  {
    "id": "nomad-cli-272",
    "nombre": "SILVIA CHEREM SACAL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MANUEL AVILA CAMACHO BLVD",
    "email": ""
  },
  {
    "id": "nomad-cli-273",
    "nombre": "SILVIA PAULINA PONCE ROGEL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JUAREZ",
    "email": ""
  },
  {
    "id": "nomad-cli-274",
    "nombre": "JUAN PABLO HERNANDEZ SEGOVIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "1ERA AVENIDA",
    "email": ""
  },
  {
    "id": "nomad-cli-275",
    "nombre": "M DE LOS ANGELES SANCHEZ JUNQUERA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JACARANDAS",
    "email": ""
  },
  {
    "id": "nomad-cli-276",
    "nombre": "OLGA SOLBES GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PASEO DEL MONARCA",
    "email": ""
  },
  {
    "id": "nomad-cli-277",
    "nombre": "ZURICH ASEGURADORA MEXICANA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BOULEVARD MANUEL AVILA CAMACHO",
    "email": ""
  },
  {
    "id": "nomad-cli-278",
    "nombre": "ERAYON SERVICIOS MEDICOS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VIALIDAD DE LA BARRANCA",
    "email": ""
  },
  {
    "id": "nomad-cli-279",
    "nombre": "MARCO ANTONIO LOPEZ MEZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "25 PONIENTE",
    "email": ""
  },
  {
    "id": "nomad-cli-280",
    "nombre": "MARIA LUISA MARTINEZ OROZCO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE 8",
    "email": ""
  },
  {
    "id": "nomad-cli-281",
    "nombre": "RAMON OLIVAS GASTELUM",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "REPUBLICA DE BOLIVIA",
    "email": ""
  },
  {
    "id": "nomad-cli-282",
    "nombre": "JOSE ARTURO BRITO MUÑOZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PRIV 3B SUR 5105 CASA 2",
    "email": ""
  },
  {
    "id": "nomad-cli-283",
    "nombre": "YRMA PATRICIA TAMAYO LEON",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CONOCIDO",
    "email": ""
  },
  {
    "id": "nomad-cli-284",
    "nombre": "HECTOR ALEJANDRO GOROSTIETA FLORES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CALLE CENTENARIO",
    "email": ""
  },
  {
    "id": "nomad-cli-285",
    "nombre": "FERNANDO MARTIN SALDIVAR GALINDO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV. CERRO GORDO",
    "email": ""
  },
  {
    "id": "nomad-cli-286",
    "nombre": "MARIA ELOISA CASILLAS MUÑOZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BOULEVARD DEL CAMPESTRE",
    "email": ""
  },
  {
    "id": "nomad-cli-287",
    "nombre": "CARLOS SANCHEZ ROMERO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CARR MEXICO VERACRUZ",
    "email": ""
  },
  {
    "id": "nomad-cli-288",
    "nombre": "SANDRA PALACIOS MARQUEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "06800",
    "email": ""
  },
  {
    "id": "nomad-cli-289",
    "nombre": "LUZ MARIA RUIZ ROSAS",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PASEO PLAYA DEL CONJCHAL",
    "email": ""
  },
  {
    "id": "nomad-cli-290",
    "nombre": "ADRIANA DE LOURDES DAVILA MORALES",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "OTE 121",
    "email": ""
  },
  {
    "id": "nomad-cli-291",
    "nombre": "ALEJANDRO LUQUE GOMEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VALLE TIZAYUCA",
    "email": ""
  },
  {
    "id": "nomad-cli-292",
    "nombre": "MARTHA ALICIA AGUIRRE ESPINOZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "BELLAS ARTES",
    "email": ""
  },
  {
    "id": "nomad-cli-293",
    "nombre": "GLORIA MARIA GARZA DE LA GARZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AV DE LAS FUENTES",
    "email": ""
  },
  {
    "id": "nomad-cli-294",
    "nombre": "MARIANA PORRAS REALZOLA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA DE LAS FUENTES",
    "email": ""
  },
  {
    "id": "nomad-cli-295",
    "nombre": "ANA LAURA NUÑEZ ARANA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "TLACOTALPAN",
    "email": ""
  },
  {
    "id": "nomad-cli-296",
    "nombre": "GLORIA MARIA TORRUCO GARZA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "JABIN",
    "email": ""
  },
  {
    "id": "nomad-cli-297",
    "nombre": "ALEJANDRO EDUARDO CASTILLO ARCE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FUENTE DE SAN MIGUEL MZ 22 LT 43",
    "email": ""
  },
  {
    "id": "nomad-cli-298",
    "nombre": "BBVA SEGUROS SALUD MEXICO, SA DE CV, GRUPO FINANCIERO BBVA MEXICO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PASEO DE LA REFORMA",
    "email": ""
  },
  {
    "id": "nomad-cli-299",
    "nombre": "FERNANDO DE ALBA QUINTANILLA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CARRANZA 31",
    "email": ""
  },
  {
    "id": "nomad-cli-300",
    "nombre": "ELSA ITZEL CERNA CERINO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "RODOLFO CAO ZAMUDIO",
    "email": ""
  },
  {
    "id": "nomad-cli-301",
    "nombre": "LORENA ORNELAS HALL",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "GELATI",
    "email": ""
  },
  {
    "id": "nomad-cli-302",
    "nombre": "LUIS ENRIQUE GUERRA GARCIA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "20 DE NOVIEMBRE",
    "email": ""
  },
  {
    "id": "nomad-cli-303",
    "nombre": "FRANCISCO MAURICIO MACEDO GUTIERREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "FAISAN",
    "email": ""
  },
  {
    "id": "nomad-cli-304",
    "nombre": "MARTA TRIGUEROS RAMIREZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "SANTISIMO",
    "email": ""
  },
  {
    "id": "nomad-cli-305",
    "nombre": "MONICA GUERRERO LEBRIJA",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "HEGEL",
    "email": ""
  },
  {
    "id": "nomad-cli-306",
    "nombre": "GERARDO ORTIZ MORENO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "RICARDO MARGAIN ZOZAYA",
    "email": ""
  },
  {
    "id": "nomad-cli-307",
    "nombre": "ERIKA ESTRADA GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "12 SUR",
    "email": ""
  },
  {
    "id": "nomad-cli-308",
    "nombre": "ROGERIO MARCELINO DOMINGUEZ MARTINEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MAR MUERTO",
    "email": ""
  },
  {
    "id": "nomad-cli-309",
    "nombre": "RODOLFO DE LEIJA GONZALEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "MORELOS",
    "email": ""
  },
  {
    "id": "nomad-cli-310",
    "nombre": "MARIA PALOMA CAMACHO LIBIEN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "VICENTE GUERRERO",
    "email": ""
  },
  {
    "id": "nomad-cli-311",
    "nombre": "JOSE ZENO EDGAR DOBERNIG GARRIDO",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "LAZARO CARDENAS 402, 11E",
    "email": ""
  },
  {
    "id": "nomad-cli-312",
    "nombre": "JAIME VON BERTRAB MESTRE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "CRR MEXICO TOLUCA",
    "email": ""
  },
  {
    "id": "nomad-cli-313",
    "nombre": "ALISTAIR MCCREADIE",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "AVENIDA CUAUHTEMOC",
    "email": ""
  },
  {
    "id": "nomad-cli-314",
    "nombre": "KARINA FUENTES MARTINEZ",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "PLATEROS",
    "email": ""
  },
  {
    "id": "nomad-cli-315",
    "nombre": "KARINA RABNER CYMERMAN",
    "rfc": "",
    "regimen": "",
    "usoCfdi": "",
    "cp": "",
    "direccion": "RETORNO",
    "email": ""
  }
];
const PRODUCTOS_BASE_SANARE = [
  {
    "id": "sanare-prd-1",
    "descripcion": "VINCRISTINA 1MG/10ML CAJ C/1 FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "VINCRI1"
  },
  {
    "id": "sanare-prd-2",
    "descripcion": "SULFATO DE MAGNESIO /G10ML FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "SULFMAG"
  },
  {
    "id": "sanare-prd-3",
    "descripcion": "SILLON DE INFUSION - DERECHO A SALA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "SILLONINFU"
  },
  {
    "id": "sanare-prd-4",
    "descripcion": "RETIRO DE INFUSOR CON HEPARINIZACIÓN",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "RET"
  },
  {
    "id": "sanare-prd-5",
    "descripcion": "PREPARACION DE MEZCLA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PREPMEZ"
  },
  {
    "id": "sanare-prd-6",
    "descripcion": "PREDNISONA 50MG 20TAB (AMSA)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PREDI50"
  },
  {
    "id": "sanare-prd-7",
    "descripcion": "PERFALGAN 1G SOLINY C/4 FCO10ML",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PERFA1G"
  },
  {
    "id": "sanare-prd-8",
    "descripcion": "PALONOSETRON (VINALTRO 025.MG/5ML AMP 1",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PALOVAINAL"
  },
  {
    "id": "sanare-prd-9",
    "descripcion": "PACLITAXEL300 MG/500ML C/1FA(SIRAPEH)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PACLITAX"
  },
  {
    "id": "sanare-prd-10",
    "descripcion": "OXALIPLATINO 50MG/10ML C/1 FA RECOPLAT",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "OXA50MG"
  },
  {
    "id": "sanare-prd-11",
    "descripcion": "OXALIPLATINO 100MG/200ML C/1 FA RECOPLAT",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "OXA100"
  },
  {
    "id": "sanare-prd-12",
    "descripcion": "ONDANSETRON ANTIVON 8MG 4MLC/1 AMP SI N",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "ONDAANT"
  },
  {
    "id": "sanare-prd-13",
    "descripcion": "OMEPRAZOL (PENTREN 40 MG AMP C/10ML)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "OMPRZ"
  },
  {
    "id": "sanare-prd-14",
    "descripcion": "METOCLOPRAMIDA 10MG/2ML CAJ C/6 AMP",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "MET1"
  },
  {
    "id": "sanare-prd-15",
    "descripcion": "MANITOL 20% 250 ML",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "MANITOL"
  },
  {
    "id": "sanare-prd-16",
    "descripcion": "MAGNEFUSIN SULFATO DE MAGNESIO PISA 10%",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "MAGNEF"
  },
  {
    "id": "sanare-prd-17",
    "descripcion": "LIDOCAINA/PRILOCAINA 1G CAJ C/2 PAR",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "LIDO/PRILO"
  },
  {
    "id": "sanare-prd-18",
    "descripcion": "KIT DE INSUMOS",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "KITINSUM"
  },
  {
    "id": "sanare-prd-19",
    "descripcion": "KETOROLACO 300 MG TB (ONEMER SUBBLINGUAL",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "KETEONESUB"
  },
  {
    "id": "sanare-prd-20",
    "descripcion": "KELEFUSIN CLORURO DE POTASIO PISA ANP 1.",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "KELEFISIN"
  },
  {
    "id": "sanare-prd-21",
    "descripcion": "INSUMOS Y SERVICIO DE INFUSION",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "SERINF"
  },
  {
    "id": "sanare-prd-22",
    "descripcion": "INFUSOR HOME PUMP",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "INFHOME"
  },
  {
    "id": "sanare-prd-23",
    "descripcion": "INFUSION GEMCITABINA+CISPLATINO",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "INFG+CIS"
  },
  {
    "id": "sanare-prd-24",
    "descripcion": "INFUSION CARBOPLATINO 3AUC+PACLITAXEL",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "INFCAR3AUC+PACLI"
  },
  {
    "id": "sanare-prd-25",
    "descripcion": "HIERRO 500MG/10ML (RENEGY)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "HIE01"
  },
  {
    "id": "sanare-prd-26",
    "descripcion": "HIDROCORTIZONA 100MG/1ML (HIDROCORTIZONA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "HIDROCORTIZONA"
  },
  {
    "id": "sanare-prd-27",
    "descripcion": "HEPARINA 100UI/1ML CAJ C/1 FA(INHEPAR)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "HEP100"
  },
  {
    "id": "sanare-prd-28",
    "descripcion": "GEMCITABINA 1G/25ML  C/1 FA(ULDEUS)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "GEMCITAB"
  },
  {
    "id": "sanare-prd-29",
    "descripcion": "GEMCITABINA (ACCOGEM 200MG/5ML F A C/)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "GEMCITAACCOGEM"
  },
  {
    "id": "sanare-prd-30",
    "descripcion": "FOSAPREPITANT 150MG/5ML CAJ C/1 FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "FOS1"
  },
  {
    "id": "sanare-prd-31",
    "descripcion": "FLUOROURACILO 250MG/10 FRA(ULSACIL)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "FLUO10"
  },
  {
    "id": "sanare-prd-32",
    "descripcion": "DOXORUBICINA 50MG/25ML CAJ C/1",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "DOXO50"
  },
  {
    "id": "sanare-prd-33",
    "descripcion": "DOCETAXEL (MIOCERKEL 20 MG  SOL INY C/1",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "DOCETAXEL"
  },
  {
    "id": "sanare-prd-34",
    "descripcion": "DIFENHIDRAMINA SOL INYECTABLE 10MG CAJ/1",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "DIFENINYEC"
  },
  {
    "id": "sanare-prd-35",
    "descripcion": "DIFENHDRAMINA SOL INYECTABLE 10 MG C/1",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "DIFENH"
  },
  {
    "id": "sanare-prd-36",
    "descripcion": "DEXAMETASONA (DECOREX 8MG/2ML CAJA C/1)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "DEXADECOREX8M2ML"
  },
  {
    "id": "sanare-prd-37",
    "descripcion": "CURACIÓN CON HEPARINIZACIÓN",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CURHEP"
  },
  {
    "id": "sanare-prd-38",
    "descripcion": "CONSULTA ONCOLÓGICA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CONSONC"
  },
  {
    "id": "sanare-prd-39",
    "descripcion": "CLORURO DE POTASIO 1.49 G/5ML CAJ/1 FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CLOPOT"
  },
  {
    "id": "sanare-prd-40",
    "descripcion": "CLOROPIRAMIDA (AVAPENA SOL. INYECTABLE)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CLORO(AVAP)"
  },
  {
    "id": "sanare-prd-41",
    "descripcion": "CISPLATINO 50MG/50ML CAJ C/1FA (ACCOCIT)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CIS50"
  },
  {
    "id": "sanare-prd-42",
    "descripcion": "CISPLATINO (ACCOCIT 10 MG/10ML FAM C/1)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CISPLA"
  },
  {
    "id": "sanare-prd-43",
    "descripcion": "CICLOFOSFAMIDA 500MG CAJ C/1 FAM(CYATA)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CICLO500"
  },
  {
    "id": "sanare-prd-44",
    "descripcion": "CICLOFOSFAMIDA 200MG/10ML CAJ C/5 FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CICLO200"
  },
  {
    "id": "sanare-prd-45",
    "descripcion": "CICLOFOSFAMIDA 1000MG CAJ C/1 FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CICLO1000"
  },
  {
    "id": "sanare-prd-46",
    "descripcion": "CARBOPLATINO450MG/45ML C/1 FA NUVAPLAST",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CARBPLAT"
  },
  {
    "id": "sanare-prd-47",
    "descripcion": "CARBOPLATINO 150MG/15ML CAJ C/1FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "CARBPLAT150"
  },
  {
    "id": "sanare-prd-48",
    "descripcion": "BLINATUMOMAB 35MCG CAJ C/1 FA(BLINCYTO)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "BLINA35"
  },
  {
    "id": "sanare-prd-49",
    "descripcion": "BEVACIZUMAB 400MG/16ML CAJ C/1",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "BEVACI400"
  },
  {
    "id": "sanare-prd-50",
    "descripcion": "BEVACIZUMAB 100MG/4ML CAJ C/1 FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "BEVACI100"
  },
  {
    "id": "sanare-prd-51",
    "descripcion": "APREPITAN (IRAMENOL 125/80MG)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "APREMIRA125/80M"
  },
  {
    "id": "sanare-prd-52",
    "descripcion": "ACIDO FOLINICO 50MG/4ML CAJ C/1 FA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "ACIFOLI"
  }
];
const PRODUCTOS_BASE_NOMAD = [
  {
    "id": "nomad-prd-1",
    "descripcion": "TEMPUS XT-XR",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "XTXR"
  },
  {
    "id": "nomad-prd-2",
    "descripcion": "TEMPUS XT +XR",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "XT+XR"
  },
  {
    "id": "nomad-prd-3",
    "descripcion": "TEMPUS XF+",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "TXF+"
  },
  {
    "id": "nomad-prd-4",
    "descripcion": "TRUSIGHT HEREDITARY CANCER PANEL",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "TRUG"
  },
  {
    "id": "nomad-prd-5",
    "descripcion": "MATERNA AVANZADA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "TRIPREM"
  },
  {
    "id": "nomad-prd-6",
    "descripcion": "MATERNA ESENCIAL",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "TRIESE"
  },
  {
    "id": "nomad-prd-7",
    "descripcion": "TEMPUS XT CDX + PDL1",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "TPSXTPD1"
  },
  {
    "id": "nomad-prd-8",
    "descripcion": "TEMPU XT CDX",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "TMPXF+"
  },
  {
    "id": "nomad-prd-9",
    "descripcion": "TIPO DE SANGRE",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "TISAN"
  },
  {
    "id": "nomad-prd-10",
    "descripcion": "TEMPUS XR + CON PDL1 CLONA 22C3",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "TEMXT+FOLR1"
  },
  {
    "id": "nomad-prd-11",
    "descripcion": "SOLID TUMOR PROFILE",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "SOTUMPROF"
  },
  {
    "id": "nomad-prd-12",
    "descripcion": "RENTA DE OFICINAS",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "RENTA"
  },
  {
    "id": "nomad-prd-13",
    "descripcion": "RECEPTOR DE FOLATOS ALFA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "RDFA"
  },
  {
    "id": "nomad-prd-14",
    "descripcion": "QYUIMICA SANGUINEA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "QUISAN"
  },
  {
    "id": "nomad-prd-15",
    "descripcion": "PROSIGNA PAM 50",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PROSPAM50"
  },
  {
    "id": "nomad-prd-16",
    "descripcion": "QIAGEN THERASCREEN PIK3CA RGQ PCR KIT",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PIK3CA"
  },
  {
    "id": "nomad-prd-17",
    "descripcion": "PAGO",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PAGO"
  },
  {
    "id": "nomad-prd-18",
    "descripcion": "PANEL DE GENOTIFICACION DE GIST",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "PAGENGIST"
  },
  {
    "id": "nomad-prd-19",
    "descripcion": "ONCOTYPE DX",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "OCTP"
  },
  {
    "id": "nomad-prd-20",
    "descripcion": "MATERNAL GEMELAR",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "MATGEM"
  },
  {
    "id": "nomad-prd-21",
    "descripcion": "ANALISIS DE PRUEBA PCR SARS-COV-2",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "MAQPCRCOV"
  },
  {
    "id": "nomad-prd-22",
    "descripcion": "EMISION DE RESULTADOS POR ANTIGENO",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "MAQANTCOV"
  },
  {
    "id": "nomad-prd-23",
    "descripcion": "MANEJO DE DEUDA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "MADEU"
  },
  {
    "id": "nomad-prd-24",
    "descripcion": "LOGISTICA DE ENVIO A FOUNDATION MEDICINE",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "LOGENV"
  },
  {
    "id": "nomad-prd-25",
    "descripcion": "KIT SOMATICO NRTK",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "KSOM"
  },
  {
    "id": "nomad-prd-26",
    "descripcion": "INVITAE",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "INV2210"
  },
  {
    "id": "nomad-prd-27",
    "descripcion": "INMUNOHISTIQUIMICA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "INMUNO"
  },
  {
    "id": "nomad-prd-28",
    "descripcion": "INESTABILIDAD MICROSATELITAL EN BIOPSIA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "INEMICRO"
  },
  {
    "id": "nomad-prd-29",
    "descripcion": "GUARDAN HEALT 360 DE 74 GENES",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "GUAHEA36074G"
  },
  {
    "id": "nomad-prd-30",
    "descripcion": "FOLR1 INMUNOHISTOQUIMICA (IHQ)",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "FOLHER"
  },
  {
    "id": "nomad-prd-31",
    "descripcion": "FOUNDATION ONE LIQUID",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "FM1 LIQUID"
  },
  {
    "id": "nomad-prd-32",
    "descripcion": "FOUNDATION ONE HEME",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "FM1 HEME"
  },
  {
    "id": "nomad-prd-33",
    "descripcion": "FOUNDATION CDX + PDL1",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "FM1 CDX+PDL1"
  },
  {
    "id": "nomad-prd-34",
    "descripcion": "FOUNDATION ONE CDX",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "FM1 CDX"
  },
  {
    "id": "nomad-prd-35",
    "descripcion": "FOUNDATION ONE LIQUID",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "FM1"
  },
  {
    "id": "nomad-prd-36",
    "descripcion": "ESTUDIO GENOMICO, METASTASIS DE PANCREAS",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "ESTGEN"
  },
  {
    "id": "nomad-prd-37",
    "descripcion": "ONCURIA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "ESONC"
  },
  {
    "id": "nomad-prd-38",
    "descripcion": "PAGO DE DEDUCIBLE",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "DEDU"
  },
  {
    "id": "nomad-prd-39",
    "descripcion": "PANBIOTM COVID-19AG RAPID TEST DEVICE",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "COVD25TEST"
  },
  {
    "id": "nomad-prd-40",
    "descripcion": "SERVICIO DE COMISION POR VENTAS",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "COMVEN"
  },
  {
    "id": "nomad-prd-41",
    "descripcion": "COMPREHENSIVE PGX  PANEL",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "COMPGX"
  },
  {
    "id": "nomad-prd-42",
    "descripcion": "MAMMAPRINT",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "COMMAMMA"
  },
  {
    "id": "nomad-prd-43",
    "descripcion": "PAGO DE COASEGURO",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "COASE"
  },
  {
    "id": "nomad-prd-44",
    "descripcion": "BRAF EN BIOPSIA LIQUIDA",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "BRAFLIQ"
  },
  {
    "id": "nomad-prd-45",
    "descripcion": "AVEIO CGP",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "AVECGP"
  },
  {
    "id": "nomad-prd-46",
    "descripcion": "PRUEBA DE ANTIGENO SARS-COV-2 TEST 5",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "ATGCOVTST5"
  },
  {
    "id": "nomad-prd-47",
    "descripcion": "PRUEBA DE ANTIGENO SARS-COV-2",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "ATGCOV"
  },
  {
    "id": "nomad-prd-48",
    "descripcion": "ARCHER VARIANTPLEX CORE MYELOID",
    "claveProdServ": "",
    "claveUnidad": "",
    "unidad": "",
    "precio": 0,
    "gravaIva": true,
    "claveInterna": "ARCVARIAN"
  }
];

// ========================= UTILIDADES =========================
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
}

function formatoMoneda(n) {
  return (n || 0).toFixed(2);
}

// Para mostrar con comas: 50,000.00
function formatoMonedaBonito(n) {
  const num = Number(n || 0);
  return num.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function escapeXml(str) {
  if (!str && str !== 0) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function hoyIso() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return (
    d.getFullYear() + "-" +
    pad(d.getMonth() + 1) + "-" +
    pad(d.getDate()) + "T" +
    pad(d.getHours()) + ":" +
    pad(d.getMinutes()) + ":" +
    pad(d.getSeconds())
  );
}



// ========================= FOLIOS POR EMPRESA =========================
function obtenerClaveFolioActual() {
  return empresaActual === "nomad" ? LS_KEYS.folioNomad : LS_KEYS.folioSanare;
}

function inicializarFolioSegunEmpresa() {
  const inputFolio = document.getElementById("fac-folio");
  if (!inputFolio) return;
  const key = obtenerClaveFolioActual();
  let ultimo = parseInt(localStorage.getItem(key) || "0", 10);
  if (!Number.isFinite(ultimo) || ultimo < 0) ultimo = 0;
  const siguiente = ultimo + 1;
  inputFolio.value = siguiente;
}

function guardarFolioUsado(folioValor) {
  const folioNum = parseInt(folioValor || "0", 10);
  if (!folioNum || folioNum <= 0) return;
  const key = obtenerClaveFolioActual();
  const actual = parseInt(localStorage.getItem(key) || "0", 10);
  if (!Number.isFinite(actual) || folioNum > actual) {
    localStorage.setItem(key, String(folioNum));
  }
}
// ========================= CARGA INICIAL =========================
document.addEventListener("DOMContentLoaded", () => {
  initNavegacion();
  cargarDesdeLocalStorage();
  cargarSatCatalogs();
  initClientes();
  initProductos();
  initCompras();
  initFacturacion();
});

// ========================= NAVEGACIÓN =========================

function initNavegacion() {
  const buttons = document.querySelectorAll(".nav-btn");
  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      buttons.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");

      const section = btn.dataset.section;
      document
        .querySelectorAll(".section")
        .forEach((sec) => sec.classList.remove("visible"));
      document
        .getElementById(`sec-${section}`)
        .classList.add("visible");
    });
  });

  const empresaButtons = document.querySelectorAll(".empresa-btn");
  empresaButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const emp = btn.dataset.empresa;
      if (emp !== "sanare" && emp !== "nomad") return;
      empresaActual = emp;
      // guardamos empresa seleccionada y recargamos catálogos
      localStorage.setItem(LS_KEYS.empresaActual, empresaActual);
      cargarDesdeLocalStorage();
      renderClientes();
      renderProductos();
      renderClientesEnFactura();
      if (typeof renderHistorialCompras === "function") {
        renderHistorialCompras();
      }
      // reiniciar folio de factura para la empresa seleccionada
      inicializarFolioSegunEmpresa();
      // actualizar estilos activos
      empresaButtons.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
    });
  });
}


// ========================= LOCAL STORAGE =========================

function cargarDesdeLocalStorage() {
  try {
    const emp = localStorage.getItem(LS_KEYS.empresaActual);
    if (emp === "sanare" || emp === "nomad") {
      empresaActual = emp;
    }

    const cliSan = localStorage.getItem(LS_KEYS.clientesSanare);
    const cliNom = localStorage.getItem(LS_KEYS.clientesNomad);
    const prdSan = localStorage.getItem(LS_KEYS.productosSanare);
    const prdNom = localStorage.getItem(LS_KEYS.productosNomad);
    const cmpSan = localStorage.getItem(LS_KEYS.comprasSanare);
    const cmpNom = localStorage.getItem(LS_KEYS.comprasNomad);

    let clientesSanare = cliSan ? JSON.parse(cliSan) : [];
    let clientesNomad = cliNom ? JSON.parse(cliNom) : [];
    let productosSanare = prdSan ? JSON.parse(prdSan) : [];
    let productosNomad = prdNom ? JSON.parse(prdNom) : [];
    let comprasSanare = cmpSan ? JSON.parse(cmpSan) : [];
    let comprasNomad = cmpNom ? JSON.parse(cmpNom) : [];

    // si están vacíos, usamos catálogo base de Excel
    if (clientesSanare.length === 0) clientesSanare = CLIENTES_BASE_SANARE.slice();
    if (clientesNomad.length === 0) clientesNomad = CLIENTES_BASE_NOMAD.slice();
    if (productosSanare.length === 0) productosSanare = PRODUCTOS_BASE_SANARE.slice();
    if (productosNomad.length === 0) productosNomad = PRODUCTOS_BASE_NOMAD.slice();

    // elegimos arreglo según empresaActual
    if (empresaActual === "sanare") {
      clientes = clientesSanare;
      productos = productosSanare;
      compras = comprasSanare;
    } else {
      clientes = clientesNomad;
      productos = productosNomad;
      compras = comprasNomad;
    }

    // guardamos de nuevo por si se inicializaron con base Excel
    localStorage.setItem(LS_KEYS.clientesSanare, JSON.stringify(clientesSanare));
    localStorage.setItem(LS_KEYS.clientesNomad, JSON.stringify(clientesNomad));
    localStorage.setItem(LS_KEYS.productosSanare, JSON.stringify(productosSanare));
    localStorage.setItem(LS_KEYS.productosNomad, JSON.stringify(productosNomad));
    localStorage.setItem(LS_KEYS.comprasSanare, JSON.stringify(comprasSanare));
    localStorage.setItem(LS_KEYS.comprasNomad, JSON.stringify(comprasNomad));
    localStorage.setItem(LS_KEYS.empresaActual, empresaActual);
  } catch (e) {
    console.warn("Error leyendo localStorage", e);
    // si algo falla, cargamos catálogos base
    clientes = empresaActual === "sanare" ? CLIENTES_BASE_SANARE.slice() : CLIENTES_BASE_NOMAD.slice();
    productos = empresaActual === "sanare" ? PRODUCTOS_BASE_SANARE.slice() : PRODUCTOS_BASE_NOMAD.slice();
    compras = [];
  }
}

function guardarEnLocalStorage() {
  // leemos lo que haya
  const cliSan = localStorage.getItem(LS_KEYS.clientesSanare);
  const cliNom = localStorage.getItem(LS_KEYS.clientesNomad);
  const prdSan = localStorage.getItem(LS_KEYS.productosSanare);
  const prdNom = localStorage.getItem(LS_KEYS.productosNomad);
  const cmpSan = localStorage.getItem(LS_KEYS.comprasSanare);
  const cmpNom = localStorage.getItem(LS_KEYS.comprasNomad);

  let clientesSanare = cliSan ? JSON.parse(cliSan) : CLIENTES_BASE_SANARE.slice();
  let clientesNomad = cliNom ? JSON.parse(cliNom) : CLIENTES_BASE_NOMAD.slice();
  let productosSanare = prdSan ? JSON.parse(prdSan) : PRODUCTOS_BASE_SANARE.slice();
  let productosNomad = prdNom ? JSON.parse(prdNom) : PRODUCTOS_BASE_NOMAD.slice();
  let comprasSanare = cmpSan ? JSON.parse(cmpSan) : [];
  let comprasNomad = cmpNom ? JSON.parse(cmpNom) : [];

  if (empresaActual === "sanare") {
    clientesSanare = clientes.slice();
    productosSanare = productos.slice();
    comprasSanare = compras.slice();
  } else {
    clientesNomad = clientes.slice();
    productosNomad = productos.slice();
    comprasNomad = compras.slice();
  }

  localStorage.setItem(LS_KEYS.clientesSanare, JSON.stringify(clientesSanare));
  localStorage.setItem(LS_KEYS.clientesNomad, JSON.stringify(clientesNomad));
  localStorage.setItem(LS_KEYS.productosSanare, JSON.stringify(productosSanare));
  localStorage.setItem(LS_KEYS.productosNomad, JSON.stringify(productosNomad));
  localStorage.setItem(LS_KEYS.comprasSanare, JSON.stringify(comprasSanare));
  localStorage.setItem(LS_KEYS.comprasNomad, JSON.stringify(comprasNomad));
  localStorage.setItem(LS_KEYS.empresaActual, empresaActual);
}


// ========================= SAT CATALOGS =========================
async function cargarSatCatalogs() {
  try {
    const res = await fetch("sat_catalogs.json");
    if (!res.ok) throw new Error("sat_catalogs.json no encontrado");
    satCatalogs = await res.json();
    console.log("SAT catalogs cargados", satCatalogs);
  } catch (e) {
    console.warn("No se pudieron cargar los catálogos SAT; uso valores mínimos", e);
    // Fallback mínimo
    satCatalogs = {
      formaPago: [
        { clave: "01", descripcion: "Efectivo" },
        { clave: "02", descripcion: "Cheque nominativo" },
        { clave: "03", descripcion: "Transferencia electrónica de fondos" }
      ],
      metodoPago: [
        { clave: "PUE", descripcion: "Pago en una sola exhibición" },
        { clave: "PPD", descripcion: "Pago en parcialidades o diferido" }
      ],
      regimenFiscal: [
        { clave: "601", descripcion: "General de Ley Personas Morales" },
        { clave: "605", descripcion: "Sueldos y Salarios e Ingresos Asimilados a Salarios" }
      ],
      usoCFDI: [
        { clave: "D01", descripcion: "Honorarios médicos, dentales y gastos hospitalarios." },
        { clave: "G03", descripcion: "Gastos en general" }
      ],
      prodServ: [],
      claveUnidad: []
    };
  }

  poblarCombosSat();
}

function poblarCombosSat() {
  // Cliente: régimen y uso
  const selReg = document.getElementById("cli-regimen");
  const selUso = document.getElementById("cli-uso-cfdi");
  selReg.innerHTML = `<option value="">Selecciona...</option>`;
  satCatalogs.regimenFiscal.forEach((r) => {
    const opt = document.createElement("option");
    opt.value = r.clave;
    opt.textContent = `${r.clave} - ${r.descripcion}`;
    selReg.appendChild(opt);
  });
  selUso.innerHTML = `<option value="">Selecciona...</option>`;
  satCatalogs.usoCFDI.forEach((u) => {
    const opt = document.createElement("option");
    opt.value = u.clave;
    opt.textContent = `${u.clave} - ${u.descripcion}`;
    selUso.appendChild(opt);
  });

  // Factura: forma de pago, método
  const selForma = document.getElementById("fac-forma-pago");
  const selMetodo = document.getElementById("fac-metodo-pago");
  selForma.innerHTML = "";
  satCatalogs.formaPago.forEach((f) => {
    const opt = document.createElement("option");
    opt.value = f.clave;
    opt.textContent = `${f.clave} - ${f.descripcion}`;
    selForma.appendChild(opt);
  });
  selMetodo.innerHTML = "";
  satCatalogs.metodoPago.forEach((m) => {
    const opt = document.createElement("option");
    opt.value = m.clave;
    opt.textContent = `${m.clave} - ${m.descripcion}`;
    selMetodo.appendChild(opt);
  });

  // Producto: datalist de ClaveProdServ y ClaveUnidad (TODOS los valores del Excel)
  const dlProdServ = document.getElementById("lista-clave-prodserv");
  const dlUnidad = document.getElementById("lista-clave-unidad");
  dlProdServ.innerHTML = "";
  satCatalogs.prodServ.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = `${p.clave} - ${p.descripcion}`;
    dlProdServ.appendChild(opt);
  });
  dlUnidad.innerHTML = "";
  satCatalogs.claveUnidad.forEach((u) => {
    const texto = u.simbolo
      ? `${u.clave} - ${u.nombre} (${u.simbolo})`
      : `${u.clave} - ${u.nombre}`;
    const opt = document.createElement("option");
    opt.value = texto;
    dlUnidad.appendChild(opt);
  });
}

// ========================= CLIENTES =========================
function initClientes() {
  const form = document.getElementById("form-cliente");
  const btnLimpiar = document.getElementById("btn-cli-limpiar");

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const id = form.dataset.editId || uid();
    const data = {
      id,
      nombre: document.getElementById("cli-nombre").value.trim(),
      rfc: document.getElementById("cli-rfc").value.trim().toUpperCase(),
      regimen: document.getElementById("cli-regimen").value,
      usoCfdi: document.getElementById("cli-uso-cfdi").value,
      cp: document.getElementById("cli-cp").value.trim(),
      direccion: document.getElementById("cli-direccion").value.trim(),
      email: document.getElementById("cli-email").value.trim(),
      telefono: document.getElementById("cli-telefono").value.trim()
    };

    if (!data.nombre || !data.rfc || !data.regimen || !data.usoCfdi || !data.cp) {
      alert("Revisa los datos obligatorios del cliente.");
      return;
    }

    const idx = clientes.findIndex((c) => c.id === id);
    if (idx >= 0) {
      clientes[idx] = data;
    } else {
      clientes.push(data);
    }
    guardarEnLocalStorage();
    renderClientes();
    renderClientesEnFactura();
    limpiarFormCliente();
  });

  btnLimpiar.addEventListener("click", limpiarFormCliente);

  renderClientes();
  renderClientesEnFactura();
}

function renderClientes() {
  const tbody = document.querySelector("#tabla-clientes tbody");
  tbody.innerHTML = "";

  clientes.forEach((c) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${c.nombre}</td>
      <td>${c.rfc}</td>
      <td>${c.regimen}</td>
      <td>${c.usoCfdi}</td>
      <td>${c.cp}</td>
      <td class="actions">
        <button type="button" data-id="${c.id}" class="btn-cli-edit">Editar</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.querySelectorAll(".btn-cli-edit").forEach((btn) => {
    btn.addEventListener("click", () => {
      editarCliente(btn.dataset.id);
    });
  });
}

function editarCliente(id) {
  const c = clientes.find((x) => x.id === id);
  if (!c) return;
  const form = document.getElementById("form-cliente");
  form.dataset.editId = c.id;
  document.getElementById("cli-nombre").value = c.nombre;
  document.getElementById("cli-rfc").value = c.rfc;
  document.getElementById("cli-regimen").value = c.regimen;
  document.getElementById("cli-uso-cfdi").value = c.usoCfdi;
  document.getElementById("cli-cp").value = c.cp;
  document.getElementById("cli-direccion").value = c.direccion;
  document.getElementById("cli-email").value = c.email;
  document.getElementById("cli-telefono").value = c.telefono;
}

function limpiarFormCliente() {
  const form = document.getElementById("form-cliente");
  form.reset();
  delete form.dataset.editId;
}

// ========================= PRODUCTOS =========================
function initProductos() {
  const form = document.getElementById("form-producto");
  const btnLimpiar = document.getElementById("btn-prd-limpiar");

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const id = form.dataset.editId || uid();

    const claveProdServText = document.getElementById("prd-clave-prodserv").value.trim();
    const claveUnidadText = document.getElementById("prd-clave-unidad").value.trim();

    const claveProdServ = claveProdServText.split(" - ")[0] || "";
    const claveUnidad = claveUnidadText.split(" - ")[0] || "";

    const data = {
      id,
      descripcion: document.getElementById("prd-descripcion").value.trim(),
      claveProdServ,
      claveUnidad,
      unidad: document.getElementById("prd-unidad").value.trim(),
      precio: parseFloat(document.getElementById("prd-precio").value || "0"),
      gravaIva: document.getElementById("prd-grava-iva").checked
    };

    if (!data.descripcion || !data.precio) {
      alert("Revisa los datos del producto.");
      return;
    }

    const idx = productos.findIndex((p) => p.id === id);
    if (idx >= 0) {
      productos[idx] = data;
    } else {
      productos.push(data);
    }
    guardarEnLocalStorage();
    renderProductos();
    limpiarFormProducto();
  });

  btnLimpiar.addEventListener("click", limpiarFormProducto);

  renderProductos();
}

function renderProductos() {
  const tbody = document.querySelector("#tabla-productos tbody");
  tbody.innerHTML = "";

  productos.forEach((p) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${p.descripcion}</td>
      <td>${p.claveProdServ}</td>
      <td>${p.claveUnidad}</td>
      <td>${p.unidad}</td>
      <td>${formatoMoneda(p.precio)}</td>
      <td>${p.gravaIva ? "Sí" : "No"}</td>
      <td class="actions">
        <button type="button" data-id="${p.id}" class="btn-prd-edit">Editar</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.querySelectorAll(".btn-prd-edit").forEach((btn) => {
    btn.addEventListener("click", () => {
      editarProducto(btn.dataset.id);
    });
  });
}

function editarProducto(id) {
  const p = productos.find((x) => x.id === id);
  if (!p) return;
  const form = document.getElementById("form-producto");
  form.dataset.editId = p.id;
  document.getElementById("prd-descripcion").value = p.descripcion;
  document.getElementById("prd-clave-prodserv").value = p.claveProdServ;
  document.getElementById("prd-clave-unidad").value = p.claveUnidad;
  document.getElementById("prd-unidad").value = p.unidad;
  document.getElementById("prd-precio").value = p.precio;
  document.getElementById("prd-grava-iva").checked = !!p.gravaIva;
}

function limpiarFormProducto() {
  const form = document.getElementById("form-producto");
  form.reset();
  delete form.dataset.editId;
}

// ========================= FACTURACIÓN =========================
function initFacturacion() {
  // Fecha actual
  document.getElementById("fac-fecha").value = hoyIso();
  inicializarFolioSegunEmpresa();

  // Cliente en combo
  renderClientesEnFactura();

  // Botones
  document
    .getElementById("btn-agregar-concepto")
    .addEventListener("click", agregarConcepto);

  document
    .getElementById("btn-generar-xml")
    .addEventListener("click", generarXml);

  document
    .getElementById("btn-ver-impresion")
    .addEventListener("click", verImpresion);

  // iniciar con una fila de conceptos
  agregarConcepto();
}

function renderClientesEnFactura() {
  const sel = document.getElementById("fac-cliente");
  sel.innerHTML = "";
  if (clientes.length === 0) {
    const opt = document.createElement("option");
    opt.textContent = "No hay clientes, crea uno primero";
    opt.value = "";
    sel.appendChild(opt);
    return;
  }
  clientes.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = `${c.nombre} (${c.rfc})`;
    sel.appendChild(opt);
  });
}

function agregarConcepto() {
  const tbody = document.getElementById("fac-conceptos-body");
  const id = uid();
  const tr = document.createElement("tr");
  tr.dataset.rowId = id;

  const opcionesProductos = productos
    .map((p) => {
      const etiqueta = p.claveInterna
        ? `${escapeXml(p.claveInterna)} - ${escapeXml(p.descripcion)}`
        : escapeXml(p.descripcion);
      return `<option value="${p.id}">${etiqueta}</option>`;
    })
    .join("");

  tr.innerHTML = `
    <td>
      <select class="fac-prod-select">
        <option value="">Selecciona...</option>
        ${opcionesProductos}
      </select>
    </td>
    <td><input type="text" class="fac-descripcion" /></td>
    <td><input type="number" min="0" step="0.0001" class="fac-cantidad" value="1" /></td>
    <td><input type="text" class="fac-clave-prodserv" list="lista-clave-prodserv" /></td>
    <td><input type="text" class="fac-clave-unidad" list="lista-clave-unidad" /></td>
    <td><input type="text" class="fac-unidad" /></td>
    <td><input type="number" min="0" step="0.000001" class="fac-precio" /></td>
    <td><input type="text" class="fac-importe" readonly /></td>
    <td><input type="number" min="0" step="0.000001" class="fac-coaseguro" list="lista-coaseguro" /></td>
    <td><input type="number" min="0" step="0.000001" class="fac-deducible" list="lista-deducible" /></td>
    <td><input type="text" class="fac-iva" readonly /></td>
    <td class="actions"><button type="button" class="fac-row-del">X</button></td>
  `;

  tbody.appendChild(tr);

  const selProd = tr.querySelector(".fac-prod-select");
  const inpDesc = tr.querySelector(".fac-descripcion");
  const inpCant = tr.querySelector(".fac-cantidad");
  const inpClaveProd = tr.querySelector(".fac-clave-prodserv");
  const inpClaveUni = tr.querySelector(".fac-clave-unidad");
const inpUni = tr.querySelector(".fac-unidad");
const inpPrecio = tr.querySelector(".fac-precio");

if (satCatalogs && Array.isArray(satCatalogs.claveUnidad)) {
  inpClaveUni.addEventListener("change", () => {
    const clave = extraerClaveSat(inpClaveUni.value);
    const match = satCatalogs.claveUnidad.find((u) => u.clave === clave);
    if (match) {
      inpUni.value = match.simbolo || match.nombre || inpUni.value;
    }
  });
}

  const inpImporte = tr.querySelector(".fac-importe");
  const inpCoaseguro = tr.querySelector(".fac-coaseguro");
  const inpDeducible = tr.querySelector(".fac-deducible");
  const inpIva = tr.querySelector(".fac-iva");
  const btnDel = tr.querySelector(".fac-row-del");

  selProd.addEventListener("change", () => {
    const p = productos.find((x) => x.id === selProd.value);
    if (!p) return;
    inpDesc.value = p.claveInterna ? `${p.claveInterna} - ${p.descripcion}` : p.descripcion;
    inpClaveProd.value = p.claveProdServ;
    inpClaveUni.value = p.claveUnidad;

    // Rellenar unidad automáticamente usando catálogos SAT si existen
    if (satCatalogs && Array.isArray(satCatalogs.claveUnidad)) {
      const clave = extraerClaveSat(p.claveUnidad || inpClaveUni.value);
      const match = satCatalogs.claveUnidad.find((u) => u.clave === clave);
      if (match) {
        inpUni.value = match.simbolo || match.nombre || p.unidad || "";
      } else {
        inpUni.value = p.unidad || "";
      }
    } else {
      inpUni.value = p.unidad || "";
    }

    inpPrecio.value = p.precio;
    calcularImportes();
  });

  [inpCant, inpPrecio].forEach((inp) => {
    inp.addEventListener("input", () => calcularImportes());
  });

  [inpCoaseguro, inpDeducible].forEach((inp) => {
    inp.addEventListener("input", () => recalcularTotales());
  });

  btnDel.addEventListener("click", () => {
    tr.remove();
    calcularImportes();
  });

  function calcularImportes() {
    const cant = parseFloat(inpCant.value || "0");
    const precio = parseFloat(inpPrecio.value || "0");
    const base = cant * precio;
    const p = productos.find((x) => x.id === selProd.value);
    const gravaIva = p ? p.gravaIva : true;
    const iva = gravaIva ? base * 0.16 : 0;
    inpImporte.value = formatoMoneda(base);
    inpIva.value = formatoMoneda(iva);
    recalcularTotales();
  }
}

function extraerClaveSat(valor) {
  const v = (valor || "").trim();
  if (!v) return "";
  const partes = v.split(" - ");
  return partes[0].trim();
}

function leerConceptosFactura() {
  const rows = Array.from(document.querySelectorAll("#fac-conceptos-body tr"));
  const conceptos = [];

  rows.forEach((tr) => {
    const cant = parseFloat(tr.querySelector(".fac-cantidad").value || "0");
    const precio = parseFloat(tr.querySelector(".fac-precio").value || "0");
    if (!cant || !precio) return;
    const base = cant * precio;
    const iva = parseFloat(tr.querySelector(".fac-iva").value || "0");
    const coaseguro = parseFloat(tr.querySelector(".fac-coaseguro")?.value || "0");
    const deducible = parseFloat(tr.querySelector(".fac-deducible")?.value || "0");

    conceptos.push({
      descripcion: tr.querySelector(".fac-descripcion").value.trim(),
      cantidad: cant,
      claveProdServ: extraerClaveSat(tr.querySelector(".fac-clave-prodserv").value),
      claveUnidad: extraerClaveSat(tr.querySelector(".fac-clave-unidad").value),
      unidad: tr.querySelector(".fac-unidad").value.trim(),
      valorUnitario: precio,
      importe: base,
      iva,
      coaseguro,
      deducible
    });
  });

  return conceptos;
}

function recalcularTotales() {
  conceptosFactura = leerConceptosFactura();
  let subtotal = 0;
  let ivaTotal = 0;
  let coaseguroTotal = 0;
  let deducibleTotal = 0;

  conceptosFactura.forEach((c) => {
    subtotal += c.importe;
    ivaTotal += c.iva;
    coaseguroTotal += c.coaseguro || 0;
    deducibleTotal += c.deducible || 0;
  });

  const totalCfdi = subtotal + ivaTotal;
  const totalPaciente = totalCfdi - coaseguroTotal - deducibleTotal;

  document.getElementById("fac-subtotal").textContent = formatoMonedaBonito(subtotal);
  document.getElementById("fac-iva").textContent = formatoMonedaBonito(ivaTotal);
  document.getElementById("fac-total").textContent = formatoMonedaBonito(totalCfdi);

  const elCoa = document.getElementById("fac-coaseguro-total");
  const elDed = document.getElementById("fac-deducible-total");
  const elTotAdj = document.getElementById("fac-total-ajustado");
  if (elCoa) elCoa.textContent = formatoMonedaBonito(coaseguroTotal);
  if (elDed) elDed.textContent = formatoMonedaBonito(deducibleTotal);
  if (elTotAdj) elTotAdj.textContent = formatoMonedaBonito(totalPaciente);
}

// ========================= GENERAR XML CFDI =========================
function generarXml() {
  conceptosFactura = leerConceptosFactura();
  if (conceptosFactura.length === 0) {
    alert("Agrega al menos un concepto.");
    return;
  }

  const clienteId = document.getElementById("fac-cliente").value;
  const cliente = clientes.find((c) => c.id === clienteId);
  if (!cliente) {
    alert("Selecciona un cliente.");
    return;
  }

  const serie = document.getElementById("fac-serie").value.trim() || "";
  const folio = document.getElementById("fac-folio").value.trim() || "";
  const formaPago = document.getElementById("fac-forma-pago").value;
  const metodoPago = document.getElementById("fac-metodo-pago").value;
  const moneda = document.getElementById("fac-moneda").value.trim() || "MXN";
  const lugarExp = document.getElementById("fac-lugar-exp").value.trim() || EMISOR.cp;
  const tipoComp = document.getElementById("fac-tipo-comp").value.trim() || "I";
  const fecha = hoyIso();
  document.getElementById("fac-fecha").value = fecha;

  let subtotal = 0;
  let ivaTotal = 0;
  conceptosFactura.forEach((c) => {
    subtotal += c.importe;
    ivaTotal += c.iva;
  });
  const total = subtotal + ivaTotal;

  const conceptosXml = conceptosFactura
    .map((c) => {
      const base = c.importe;
      const importeIva = c.iva;
      return `
  <cfdi:Concepto ObjetoImp="02"
    ClaveProdServ="${escapeXml(c.claveProdServ || "01010101")}"
    Cantidad="${c.cantidad.toFixed(4)}"
    ClaveUnidad="${escapeXml(c.claveUnidad || "E48")}"
    Unidad="${escapeXml(c.unidad || "SERV")}"
    Descripcion="${escapeXml(c.descripcion)}"
    ValorUnitario="${c.valorUnitario.toFixed(6)}"
    Importe="${base.toFixed(2)}">
    <cfdi:Impuestos>
      <cfdi:Traslados>
        <cfdi:Traslado Base="${base.toFixed(2)}"
          Impuesto="002"
          TipoFactor="Tasa"
          TasaOCuota="0.160000"
          Importe="${importeIva.toFixed(2)}" />
      </cfdi:Traslados>
    </cfdi:Impuestos>
  </cfdi:Concepto>`;
    })
    .join("");

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante
  Version="4.0"
  Serie="${escapeXml(serie)}"
  Folio="${escapeXml(folio)}"
  Fecha="${fecha}"
  Moneda="${escapeXml(moneda)}"
  Exportacion="01"
  TipoDeComprobante="${escapeXml(tipoComp)}"
  FormaPago="${escapeXml(formaPago)}"
  MetodoPago="${escapeXml(metodoPago)}"
  LugarExpedicion="${escapeXml(lugarExp)}"
  SubTotal="${subtotal.toFixed(2)}"
  Total="${total.toFixed(2)}"
  xmlns:cfdi="http://www.sat.gob.mx/cfd/4"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="http://www.sat.gob.mx/cfd/4 http://www.sat.gob.mx/sitio_internet/cfd/4/cfdv40.xsd">
  <cfdi:Emisor
    Rfc="${escapeXml(EMISOR.rfc)}"
    Nombre="${escapeXml(EMISOR.nombre)}"
    RegimenFiscal="${escapeXml(EMISOR.regimenFiscal)}" />
  <cfdi:Receptor
    Rfc="${escapeXml(cliente.rfc)}"
    Nombre="${escapeXml(cliente.nombre)}"
    DomicilioFiscalReceptor="${escapeXml(cliente.cp)}"
    RegimenFiscalReceptor="${escapeXml(cliente.regimen)}"
    UsoCFDI="${escapeXml(cliente.usoCfdi)}" />
  <cfdi:Conceptos>${conceptosXml}
  </cfdi:Conceptos>
  <cfdi:Impuestos TotalImpuestosTrasladados="${ivaTotal.toFixed(2)}">
    <cfdi:Traslados>
      <cfdi:Traslado Base="${subtotal.toFixed(2)}"
        Impuesto="002"
        TipoFactor="Tasa"
        TasaOCuota="0.160000"
        Importe="${ivaTotal.toFixed(2)}" />
    </cfdi:Traslados>
  </cfdi:Impuestos>
  <!-- Complemento de timbrado se agrega cuando el PAC timbra este CFDI -->
</cfdi:Comprobante>`;

  document.getElementById("xml-output").value = xml;
  actualizarVistaImpresion(cliente, { serie, folio, fecha, formaPago, metodoPago, subtotal, ivaTotal, total });
  guardarFolioUsado(folio);
}

// ========================= VISTA IMPRESIÓN =========================
function actualizarVistaImpresion(cliente, datosFactura) {
  const preview = document.getElementById("preview-impresion");
  const conceptosHtml = conceptosFactura
    .map((c, i) => {
      return `
      <tr>
        <td>${i + 1}</td>
        <td>${escapeXml(c.descripcion)}</td>
        <td>${c.cantidad.toFixed(2)}</td>
        <td>${escapeXml(c.unidad || "")}</td>
        <td>${formatoMonedaBonito(c.valorUnitario)}</td>
        <td>${formatoMonedaBonito(c.importe)}</td>
        <td>${formatoMonedaBonito(c.coaseguro || 0)}</td>
        <td>${formatoMonedaBonito(c.deducible || 0)}</td>
        <td>${formatoMonedaBonito(c.iva || 0)}</td>
      </tr>`;
    })
    .join("");

    const logoHtml =
    empresaActual === "nomad"
      ? `<div><img src="logo_nomad.png" alt="Nomad Genetics" style="max-height:60px;margin-bottom:4px;" /></div>`
      : empresaActual === "sanare"
      ? `<div><img src="logo_sanare.png" alt="Sanaré Clínica de Infusión" style="max-height:60px;margin-bottom:4px;" /></div>`
      : "";

const html = `
  <div class="factura-pdf">
    <div class="factura-pdf-header">
      <div class="emisor">
        ${logoHtml}
        <h2>Datos del cliente</h2>
        <div class="factura-pdf-datos">
          <strong>${escapeXml(EMISOR.nombre)}</strong><br/>
          RFC: ${escapeXml(EMISOR.rfc)}<br/>
          ${escapeXml(EMISOR.direccion)}<br/>
          Régimen fiscal: ${escapeXml(EMISOR.regimenFiscal)}<br/>
          Lugar de expedición: ${escapeXml(EMISOR.cp)}
        </div>
      </div>
      <div class="receptor">
        <h2>Receptor</h2>
        <div class="factura-pdf-datos">
          <strong>${escapeXml(cliente.nombre)}</strong><br/>
          RFC: ${escapeXml(cliente.rfc)}<br/>
          CP: ${escapeXml(cliente.cp)}<br/>
          Uso CFDI: ${escapeXml(cliente.usoCfdi)}<br/>
          Régimen fiscal: ${escapeXml(cliente.regimen)}<br/>
          Dirección: ${escapeXml(cliente.direccion)}
        </div>
      </div>
    </div>

    <div class="factura-pdf-datos" style="margin-top:6px;">
      Comprobante fiscal digital - Serie: ${escapeXml(datosFactura.serie)} Folio: ${escapeXml(datosFactura.folio)}<br/>
      Fecha y hora: ${escapeXml(datosFactura.fecha)}<br/>
      Forma de pago: ${escapeXml(datosFactura.formaPago)} -
      Método de pago: ${escapeXml(datosFactura.metodoPago)}<br/>
      Moneda: MXN - Tipo de comprobante: I
    </div>

    <div class="factura-pdf-table">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Descripción</th>
            <th>Cantidad</th>
            <th>Unidad</th>
            <th>P/U</th>
            <th>Importe</th>
            <th>Coaseguro</th>
            <th>Deducible</th>
            <th>IVA</th>
          </tr>
        </thead>
        <tbody>
          ${conceptosHtml}
        </tbody>
      </table>
    </div>

    <div class="factura-pdf-totales">
      Subtotal: ${formatoMonedaBonito(datosFactura.subtotal)}<br/>
      IVA 16%: ${formatoMonedaBonito(datosFactura.ivaTotal)}<br/>
      Total coaseguro: ${formatoMonedaBonito(datosFactura.coaseguroTotal || 0)}<br/>
      Total deducible: ${formatoMonedaBonito(datosFactura.deducibleTotal || 0)}<br/>
      Total CFDI: ${formatoMonedaBonito(datosFactura.total)}<br/>
      <strong>Total paciente: ${formatoMonedaBonito(datosFactura.totalPaciente != null ? datosFactura.totalPaciente : datosFactura.total)}</strong>
    </div>

    <div class="factura-pdf-sellos">
      <strong>Sellos y cadena original</strong><br/>
      (Aquí puedes imprimir el sello del CFDI, sello SAT y cadena original una vez timbrado)
    </div>
  </div>`;

  preview.innerHTML = html;

  // también para el área de impresión
  document.getElementById("print-area").innerHTML = html;
}

function verImpresion() {
  const clienteId = document.getElementById("fac-cliente").value;
  const cliente = clientes.find((c) => c.id === clienteId);
  if (!cliente) {
    alert("Selecciona un cliente antes de imprimir.");
    return;
  }
  conceptosFactura = leerConceptosFactura();
  let subtotal = 0;
  let ivaTotal = 0;
  let coaseguroTotal = 0;
  let deducibleTotal = 0;
  conceptosFactura.forEach((c) => {
    subtotal += c.importe;
    ivaTotal += c.iva;
    coaseguroTotal += c.coaseguro || 0;
    deducibleTotal += c.deducible || 0;
  });
  const totalCfdi = subtotal + ivaTotal;
  const totalPaciente = totalCfdi - coaseguroTotal - deducibleTotal;
  const datosFactura = {
    serie: document.getElementById("fac-serie").value.trim(),
    folio: document.getElementById("fac-folio").value.trim(),
    fecha: document.getElementById("fac-fecha").value.trim() || hoyIso(),
    formaPago: document.getElementById("fac-forma-pago").value,
    metodoPago: document.getElementById("fac-metodo-pago").value,
    subtotal,
    ivaTotal,
    total: totalCfdi,
    coaseguroTotal,
    deducibleTotal,
    totalPaciente
  };
  actualizarVistaImpresion(cliente, datosFactura);
  guardarFolioUsado(datosFactura.folio);
  // Pequeña espera para que las imágenes (logos) terminen de cargarse antes de imprimir
  setTimeout(() => {
    window.print();
  }, 500);
}


// ========================= COMPRAS (ESTILO SAE) =========================
function initCompras() {
  const tipoEl = document.getElementById("comp-tipo");
  if (!tipoEl) return;

  document
    .getElementById("btn-agregar-partida")
    .addEventListener("click", agregarPartidaCompra);

  document
    .getElementById("btn-nueva-compra")
    .addEventListener("click", () => nuevaCompra(true));

  document
    .getElementById("btn-guardar-compra")
    .addEventListener("click", guardarCompraActual);

  document
    .getElementById("compras-detalle-body")
    .addEventListener("input", (e) => {
      if (e.target.classList.contains("cmp-cant") ||
          e.target.classList.contains("cmp-costo")) {
        recalcularSubtotalesCompras();
      }
    });

  document
    .getElementById("compras-detalle-body")
    .addEventListener("click", (e) => {
      if (e.target.classList.contains("btn-del-partida")) {
        const tr = e.target.closest("tr");
        tr.remove();
        recalcularSubtotalesCompras();
      }
    });

  // cargar compras de esta empresa
  cargarComprasDesdeLocalStorage();
  nuevaCompra(false);
  renderHistorialCompras();
}

function cargarComprasDesdeLocalStorage() {
  try {
    const key = empresaActual === "sanare" ? LS_KEYS.comprasSanare : LS_KEYS.comprasNomad;
    const str = localStorage.getItem(key);
    compras = str ? JSON.parse(str) : [];
  } catch (e) {
    console.warn("Error leyendo compras de localStorage", e);
    compras = [];
  }
}

function guardarComprasEnLocalStorage() {
  try {
    const key = empresaActual === "sanare" ? LS_KEYS.comprasSanare : LS_KEYS.comprasNomad;
    localStorage.setItem(key, JSON.stringify(compras));
  } catch (e) {
    console.warn("No se pudieron guardar las compras", e);
  }
}

function nuevaCompra(incrementarConsecutivo) {
  const numeroInput = document.getElementById("comp-numero");

  let nuevoNumero = 1;
  if (compras.length) {
    const nums = compras
      .map((c) => parseInt(c.numero || "0", 10) || 0);
    const max = Math.max.apply(null, nums);
    nuevoNumero = max + 1;
  }

  if (!incrementarConsecutivo && numeroInput.value) {
    const limpio = numeroInput.value.replace(/^0+/, "") || "1";
    nuevoNumero = parseInt(limpio, 10);
  }

  numeroInput.value = nuevoNumero.toString().padStart(8, "0");

  const ids = [
    "comp-fecha",
    "comp-proveedor",
    "comp-rfc",
    "comp-nombre",
    "comp-ref-prov",
    "comp-esquema",
    "comp-calle",
    "comp-num-ext",
    "comp-num-int",
    "comp-colonia",
    "comp-cp",
    "comp-poblacion",
    "comp-pais",
    "comp-entregar-a",
    "comp-fecha-rec",
    "comp-descuento",
    "comp-desc-fin",
    "comp-gastos-ind",
    "comp-almacen",
    "comp-descripcion",
    "comp-fact-unidades",
    "comp-total-partida"
  ];
  ids.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });

  const hoy = new Date().toISOString().slice(0, 10);
  const fechaEl = document.getElementById("comp-fecha");
  if (fechaEl && !fechaEl.value) fechaEl.value = hoy;
  const fechaRecEl = document.getElementById("comp-fecha-rec");
  if (fechaRecEl && !fechaRecEl.value) fechaRecEl.value = hoy;

  const tbody = document.getElementById("compras-detalle-body");
  tbody.innerHTML = "";
  agregarPartidaCompra();
  recalcularSubtotalesCompras();
}

function agregarPartidaCompra() {
  const tbody = document.getElementById("compras-detalle-body");
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input type="number" min="0" step="0.01" class="cmp-cant" value="1" /></td>
    <td><input type="text" class="cmp-prod" /></td>
    <td><input type="text" class="cmp-unidad" /></td>
    <td><input type="number" step="0.01" class="cmp-desc" /></td>
    <td><input type="number" step="0.01" class="cmp-ieps" /></td>
    <td><input type="number" step="0.01" class="cmp-coaseguro" /></td>
    <td><input type="number" step="0.01" class="cmp-deducible" /></td>
    <td><input type="number" step="0.01" class="cmp-iva" /></td>
    <td><input type="number" min="0" step="0.01" class="cmp-costo" /></td>
    <td class="compras-total-sub">0.00</td>
    <td><button type="button" class="btn-del-partida">✕</button></td>
  `;
  tbody.appendChild(tr);
  recalcularSubtotalesCompras();
}

function recalcularSubtotalesCompras() {
  const tbody = document.getElementById("compras-detalle-body");
  const filas = Array.from(tbody.querySelectorAll("tr"));
  let total = 0;
  filas.forEach((tr) => {
    const cant = parseFloat(tr.querySelector(".cmp-cant")?.value || "0");
    const costo = parseFloat(tr.querySelector(".cmp-costo")?.value || "0");
    const subtotal = Math.round((cant * costo + Number.EPSILON) * 100) / 100;
    total += subtotal;
    const celdaSub = tr.querySelector(".compras-total-sub");
    if (celdaSub) celdaSub.textContent = subtotal.toFixed(2);
  });
  const totalInput = document.getElementById("comp-total-partida");
  if (totalInput) totalInput.value = total.toFixed(2);
}

function guardarCompraActual() {
  const numero = document.getElementById("comp-numero").value;
  const tipo = document.getElementById("comp-tipo").value;
  const fecha = document.getElementById("comp-fecha").value || new Date().toISOString().slice(0, 10);
  const proveedor = document.getElementById("comp-proveedor").value;
  const descripcion = document.getElementById("comp-descripcion").value;
  const total = parseFloat(document.getElementById("comp-total-partida").value || "0");

  const tbody = document.getElementById("compras-detalle-body");
  const filas = Array.from(tbody.querySelectorAll("tr"));

  if (!filas.length) {
    alert("Agrega al menos una partida.");
    return;
  }

  const partidas = filas.map((tr) => ({
    cantidad: parseFloat(tr.querySelector(".cmp-cant")?.value || "0"),
    producto: tr.querySelector(".cmp-prod")?.value || "",
    unidad: tr.querySelector(".cmp-unidad")?.value || "",
    descuento: tr.querySelector(".cmp-desc")?.value || "",
    ieps: tr.querySelector(".cmp-ieps")?.value || "",
    coaseguro: tr.querySelector(".cmp-coaseguro")?.value || "",
    deducible: tr.querySelector(".cmp-deducible")?.value || "",
    iva: tr.querySelector(".cmp-iva")?.value || "",
    costo: tr.querySelector(".cmp-costo")?.value || "",
    subtotal: parseFloat(tr.querySelector(".compras-total-sub")?.textContent || "0")
  }));

  const compra = {
    numero,
    tipo,
    fecha,
    proveedor,
    descripcion,
    total,
    partidas
  };

  const idx = compras.findIndex((c) => c.numero === numero);
  if (idx >= 0) {
    compras[idx] = compra;
  } else {
    compras.push(compra);
  }

  guardarComprasEnLocalStorage();
  renderHistorialCompras();
  alert("Compra guardada.");
}

function renderHistorialCompras() {
  const tbody = document.querySelector("#tabla-compras tbody");
  if (!tbody) return;
  tbody.innerHTML = "";
  compras
    .slice()
    .sort((a, b) => (a.numero || "").localeCompare(b.numero || ""))
    .forEach((c) => {
      const tr = document.createElement("tr");
      const concepto = c.descripcion || (c.partidas && c.partidas[0]?.producto) || "";
      tr.innerHTML = `
        <td>${c.fecha || ""}</td>
        <td>${c.proveedor || ""}</td>
        <td>${concepto}</td>
        <td style="text-align:right;">${(c.total || 0).toFixed(2)}</td>
      `;
      tbody.appendChild(tr);
    });
}
