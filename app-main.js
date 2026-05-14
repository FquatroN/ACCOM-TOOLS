const SHEET_NAME = "Comunicações";
const REVIEW_LIST_PAGE_SIZE = 100;
const REVIEW_FETCH_PAGE_SIZE = 1500;
const AUTO_REFRESH_INTERVAL_MS = 6 * 60 * 1000;
const AUTO_REFRESH_ON_FOCUS_AFTER_MS = 120 * 1000;
const REVIEW_ANALYSIS_COLORS = ["#0d6e6e", "#d97706", "#2563eb", "#b4235a", "#6d5dfc", "#16803a", "#7c3aed", "#c2410c"];
const REVIEW_SUBSCORE_KEYS = ["staff", "cleanliness", "location", "facilities", "comfort", "value_for_money"];
const DEFAULT_REVIEW_SOURCES = [
  { key: "booking", label: "Booking.com", active: true },
  { key: "hostelworld", label: "Hostelworld", active: true },
  { key: "expedia", label: "Expedia", active: true },
  { key: "airbnb", label: "Airbnb", active: true },
  { key: "vrbo", label: "VRBO", active: true },
  { key: "tripadvisor", label: "Tripadvisor", active: true },
  { key: "google", label: "Google", active: true },
];

const LOST_FOUND_STORED_OPTIONS = ["Receção", "Arrecadação 21"];
const LOST_FOUND_NUMBER_OFFSET = 8719;
const APP_FEATURE_OPTIONS = ["communications", "guests", "cash", "lost-found", "reviews", "groups", "services", "shopping", "hours", "bakery", "laundry"];
const SETTINGS_FEATURE_OPTIONS = ["general", "communications", "guests", "cash", "reviews", "groups", "services", "shopping", "hours", "bakery", "laundry", "admin-users"];
const SHOPPING_CATEGORY_OPTIONS = ["Breakfast", "Cleaning", "Sales", "Activities", "Other", "Tapas", "Utensils"];
const SHOPPING_STORED_OPTIONS = [
  "20 (10) -Frigorificos",
  "11-Armario",
  "11-Escritorio",
  "20-Lavandaria",
  "20-Limpeza",
  "21-Comidas",
  "146-Arrecadacao",
];
const DEFAULT_SHOPPING_CATEGORY_COLORS = {
  Breakfast: "#93CDDD",
  Cleaning: "#C3D69B",
  Sales: "#10253F",
  Activities: "#B3A2C7",
  Other: "#FAC090",
  Tapas: "#77933C",
  Utensils: "#D99694",
};
const SHOPPING_WEEKDAY_OPTIONS = [
  { key: "monday", label: "Monday" },
  { key: "tuesday", label: "Tuesday" },
  { key: "wednesday", label: "Wednesday" },
  { key: "thursday", label: "Thursday" },
  { key: "friday", label: "Friday" },
  { key: "saturday", label: "Saturday" },
  { key: "sunday", label: "Sunday" },
];

const DEFAULT_SETTINGS = {
  general: {
    emailConfig: {
      provider: "resend",
      smtpHost: "smtp.gmail.com",
      smtpPort: 465,
      smtpSecure: true,
      smtpUser: "",
      smtpPassword: "",
      fromEmail: "",
      fromName: "Lisboa Central Hostel",
    },
  },
  communications: {
    categories: [
      { name: "Warning", color: "#ffd89b" },
      { name: "Maintenance", color: "#a9f0df" },
      { name: "Information", color: "#add4ff" },
      { name: "Task", color: "#ffb3c2" },
    ],
    emailAutomation: {
      enabled: false,
      frequency: "everyday",
      timeOfDay: "00:00",
      recipients: [],
      frequency2: "everyday",
      timeOfDay2: "00:00",
      recipients2: [],
    },
  },
};

const DEFAULT_GROUP_ROOM_TYPES = [
  ["11 Bed Dorm Shared Bathroom", "105"],
  ["10 Bed Dorm Shared Bathroom", "102, 105"],
  ["9 Bed Dorm Shared Bathroom", "102, 105, 206"],
  ["8 Bed Dorm Shared Bathroom", "206, 203, 113, 217, 102, 105"],
  ["7 Bed Dorm Shared Bathroom", "206, 203, 113, 217, 213"],
  ["6 Bed Dorm Shared Bathroom", "206, 203, 113, 217, 213"],
  ["5 Bed Dorm Shared Bathroom", "206, 203, 113, 217, 201, 211, 213, 111"],
  ["4 Bed Dorm Shared Bathroom", "201, 202, 211, 213, 111"],
  ["4 Bed Dorm Private Bathroom", "213, 111"],
  ["3 Bed Dorm Shared Bathroom", "201, 202, 211, 213, 111, 205, 216, 212, 204, 214"],
  ["3 Bed Dorm Private Bathroom", "213, 111, 205, 204, 214"],
  ["2 Bed Dorm Shared Bathroom", "204, 205, 218, 212, 214, 215, 216, 218"],
  ["2 Bed Dorm Private Bathroom", "204, 205, 214, 215"],
  ["Twin Private with Private Bathroom", "204, 205, 214, 215"],
  ["Twin Private with Shared Bathroom", "204, 205, 218, 212, 214, 215, 216, 218"],
  ["Single Private with Private Bathroom", "204, 205, 214, 215, 112"],
  ["Single Private with Shared Bathroom", "204, 205, 218, 212, 214, 215, 216, 218, 114, 112"],
].map(([name, rooms]) => ({ name, guestsPerRoom: inferGuestsPerGroupRoomType(name), rooms: rooms.split(",").map((room) => room.trim()) }));

const DEFAULT_GROUP_SETTINGS = {
  depositPercentage: 30,
  lastPaymentDaysBeforeArrival: 14,
  emailTemplate: `Dear {{name}},

Thank you for contacting us.
Please find below our proposal based on your request:

Arrival: {{arrival}}
Departure: {{departure}} ({{nights}} nights)

{{room_table}}

Accommodation Total = {{accommodation_total}}

City Tax {{guests}} guests x {{city_tax_nights}} nights x 4€ = {{city_tax_total}}

Total = {{total}}

The price includes: bed sheets, fully equipped kitchen, 24h reception, free internet (computers in the lobby and Wi-Fi throughout the entire hostel) plus lots of information about Lisbon.
We also offer a free breakfast served daily from 8:00 AM to 11:00 AM, which includes a generous variety of options: three types of cereals, three types of bread, muffins, mini croissants, jam, honey, butter, peanut butter, chocolate cream, fruit, coffee, tea, cocoa, milk, juice and our homemade pancakes!

A {{deposit_percentage}}% non-refundable deposit ({{deposit_value}}) is required to confirm the reservation. The remaining balance must be paid up to {{last_payment_days}} days before arrival.

Payment can be made by:
Bank transfer
Credit card (we can send a secure payment link)

Bank details:
IBAN: PT50 0035 0137 00004852230 14
BIC/SWIFT: CGDIPTPL

We can also provide some tours and activities like Lisbon Walking Tours, Surf Lessons, PubCrawls. Please let us know if you need further information.

Cancelation Policy: {{deposit_percentage}}% after booking confirmation, 100% if canceled less than {{last_payment_days}} days before check-in.

Please note that there is a city tax of EUR 4 per person, per night that applies to all guests aged 13 and older. The amount of this TAX is already in the price total above. It is subject to a maximum amount of EUR 28 per guest.

Please let us know if you need any additional information.

Hope to hear from you soon,`,
  confirmationTemplate: `Dear {{name}},

Thank you for your contact.
Your reservation has been confirmed as follows:

{{confirmation_table}}

To make any changes to an existing reservation, please contact us.
Please also let us know your expected arrival time.
Please note that any cancellations must be notified at least {{last_payment_days}} days in advance (only full rooms are accepted), otherwise the total of the reservation will be charged.

Our bank details are:
IBAN: PT50 0035 0137 00004852230 14
BIC SWIFT: CGDIPTPL

Best regards,`,
  finalConfirmationTemplate: `Dear {{name}},

Thank you for your payment.
Your reservation is now fully paid and confirmed as follows:

{{confirmation_table}}

To make any changes to an existing reservation, please contact us.
Please also let us know your expected arrival time.

Best regards,`,
  roomTypes: clone(DEFAULT_GROUP_ROOM_TYPES),
};

const DEFAULT_SERVICE_PRICE_MATRIX = {
  oneWay: {
    "1-3": 35,
    "4-7": 55,
    "8-11": 90,
    "12-16": 110,
  },
  returnTrip: {
    "1-3": 63,
    "4-7": 99,
    "8-11": 162,
    "12-16": 198,
  },
};

function defaultServiceConfirmationTemplate(serviceType = "Service", airportTransfer = false) {
  const intro = airportTransfer
    ? "Your transfer is confirmed with the following details:"
    : "Your service is confirmed with the following details:";
  const airportParagraph = airportTransfer
    ? "\nFor pick-up at the airport, the transfer company will be waiting for you at arrivals with your name on a board. The pickup time is based on the flight arrival time and the transfer company will track your flight."
    : "";
  const pickupParagraph = "\nFor pick-up in other locations, please be ready 5 minutes before the scheduled pickup time.";
  const paymentParagraph = airportTransfer
    ? "\nPayment should be made at the check-in desk, not to the driver."
    : "";
  return `Dear {{customer_name}},

${intro}

{{service_table}}${airportParagraph}
${pickupParagraph}

If you have any trouble, please use the shuttle service number: +351 917921578. It is also available on WhatsApp. Please contact the company if you have any problem finding them, otherwise we will have to charge the service amount.${paymentParagraph}

Cancellation Policy

Any cancellations must be informed 48h before service, otherwise the full amount of the service will be charged.

Best regards,
Lisboa Central Hostel`;
}

const SERVICE_CONFIRMATION_TEMPLATES = {
  pt: (airportTransfer = false) => `Caro/a {{customer_name}},

O seu serviço está confirmado com os seguintes detalhes:

{{service_table}}

${airportTransfer
    ? "Para recolhas no aeroporto, a empresa de transfer estará à sua espera nas chegadas com o seu nome numa placa. A hora da recolha baseia-se na hora de chegada do voo e a empresa acompanhará o voo."
    : ""}

Por favor esteja pronto/a 5 minutos antes da hora marcada para a recolha.

Se tiver alguma dificuldade, por favor utilize o número da empresa de transfer: +351 917921578. Também está disponível no WhatsApp. Se tiver dificuldade em encontrar a empresa, contacte-a diretamente; caso contrário, teremos de cobrar o valor do serviço.

${airportTransfer ? "O pagamento deve ser efetuado na receção no momento do check-in, e não ao motorista.\n\n" : ""}Política de cancelamento

Qualquer cancelamento deve ser informado com 48h de antecedência, caso contrário será cobrado o valor total do serviço.

Com os melhores cumprimentos,
Lisboa Central Hostel`,
  es: (airportTransfer = false) => `Estimado/a {{customer_name}},

Su servicio está confirmado con los siguientes detalles:

{{service_table}}

${airportTransfer
    ? "Para recogidas en el aeropuerto, la empresa de traslado estará esperándole en llegadas con su nombre en un cartel. La hora de recogida se basa en la hora de llegada del vuelo y la empresa hará el seguimiento del vuelo."
    : ""}

Por favor esté listo/a 5 minutos antes de la hora programada para la recogida.

Si tiene algún problema, por favor utilice el número de la empresa de traslado: +351 917921578. También está disponible en WhatsApp. Si tiene alguna dificultad para encontrar a la empresa, póngase en contacto directamente con ella; de lo contrario, tendremos que cobrar el importe del servicio.

${airportTransfer ? "El pago debe realizarse en la recepción durante el check-in, no al conductor.\n\n" : ""}Política de cancelación

Cualquier cancelación debe comunicarse con 48h de antelación; de lo contrario, se cobrará el importe total del servicio.

Saludos cordiales,
Lisboa Central Hostel`,
};

const DEFAULT_SERVICE_SETTINGS = {
  automaticEmailRecipients: [],
  liveFlightStatusEnabled: true,
  serviceConfigs: [
    {
      id: "airport-transfer",
      serviceType: "Airport Transfer",
      providerUserId: "",
      providerEmail: "odete@netcabo.pt",
      airportTransfer: true,
      hasReturn: true,
      approvedByDefault: false,
      priceMode: "airport_matrix",
      priceMatrix: clone(DEFAULT_SERVICE_PRICE_MATRIX),
      confirmationTemplate: defaultServiceConfirmationTemplate("Airport Transfer", true),
    },
    {
      id: "other-transfer",
      serviceType: "Other Transfer",
      providerUserId: "",
      providerEmail: "odete@netcabo.pt",
      airportTransfer: true,
      hasReturn: true,
      approvedByDefault: false,
      priceMode: "open",
      priceMatrix: { oneWay: {}, returnTrip: {} },
      confirmationTemplate: defaultServiceConfirmationTemplate("Other Transfer", true),
    },
    {
      id: "tour",
      serviceType: "Tour",
      providerUserId: "",
      providerEmail: "",
      airportTransfer: false,
      hasReturn: false,
      approvedByDefault: false,
      priceMode: "open",
      priceMatrix: { oneWay: {}, returnTrip: {} },
      confirmationTemplate: defaultServiceConfirmationTemplate("Tour", false),
    },
    {
      id: "boat-tour",
      serviceType: "Boat Tour",
      providerUserId: "",
      providerEmail: "",
      airportTransfer: false,
      hasReturn: false,
      approvedByDefault: false,
      priceMode: "open",
      priceMatrix: { oneWay: {}, returnTrip: {} },
      confirmationTemplate: defaultServiceConfirmationTemplate("Boat Tour", false),
    },
  ],
};

const DEFAULT_SHOPPING_SETTINGS = {
  mandatoryWeekdays: [],
  emailRecipients: [],
  categoryColors: { ...DEFAULT_SHOPPING_CATEGORY_COLORS },
  items: [],
};

const CASH_DENOMINATIONS = [
  { key: "500", value: 500 },
  { key: "200", value: 200 },
  { key: "100", value: 100 },
  { key: "50", value: 50 },
  { key: "20", value: 20 },
  { key: "10", value: 10 },
  { key: "5", value: 5 },
  { key: "2", value: 2 },
  { key: "1", value: 1 },
  { key: "0.5", value: 0.5 },
  { key: "0.2", value: 0.2 },
  { key: "0.1", value: 0.1 },
  { key: "0.05", value: 0.05 },
  { key: "0.02", value: 0.02 },
  { key: "0.01", value: 0.01 },
];
const CASH_MIN_ALERT_DENOMINATIONS = ["500", "200", "50", "20", "10", "5", "2", "1", "0.5", "0.2", "0.1"];

const DEFAULT_CASH_SETTINGS = {
  shifts: [
    { id: "night", name: "Night", startTime: "00:00" },
    { id: "morning", name: "Morning", startTime: "08:00" },
    { id: "afternoon", name: "Afternoon", startTime: "16:00" },
  ],
  items: [
    { id: "chaves-2d", name: "Chaves 2D", defaultQuantity: 6 },
    { id: "chaves-2e", name: "Chaves 2E", defaultQuantity: 6 },
    { id: "chaves-3e", name: "Chaves 3E", defaultQuantity: 6 },
    { id: "chaves-4d", name: "Chaves 4D", defaultQuantity: 6 },
    { id: "chaves-4e", name: "Chaves 4E", defaultQuantity: 6 },
    { id: "chaves-5e", name: "Chaves 5E", defaultQuantity: 6 },
    { id: "chaves-5d", name: "Chaves 5D", defaultQuantity: 6 },
    { id: "arrecadacao-2d", name: "Arrecadacao 2D", defaultQuantity: 6 },
    { id: "arrecadacao-exterior", name: "Arrecadacao Exterior", defaultQuantity: 3 },
    { id: "camas-extras-2e", name: "Camas Extras 2E", defaultQuantity: 2 },
    { id: "arrecadacao-2-5", name: "Arrecadacao 2º (5)", defaultQuantity: 5 },
    { id: "arrecadacao-3-e-1", name: "Arrecadacao 3ºE (1)", defaultQuantity: 1 },
    { id: "camas-extra-4-e-3", name: "Camas Extra 4ºE (3)", defaultQuantity: 3 },
    { id: "arrecadacao-4-d-2", name: "Arrecadacao 4ºD (2)", defaultQuantity: 2 },
    { id: "armario-5-d-1", name: "Armario 5ºD (1)", defaultQuantity: 1 },
    { id: "armario-5-e-1", name: "Armario 5ºE (1)", defaultQuantity: 1 },
    { id: "secadores", name: "Secadores", defaultQuantity: 3 },
    { id: "comandos", name: "Comandos", defaultQuantity: 7 },
    { id: "chaves-staff", name: "Chaves Staff", defaultQuantity: 8 },
  ],
  minCash: {
    "500": 0,
    "200": 0,
    "50": 0,
    "20": 0,
    "10": 0,
    "5": 0,
    "2": 0,
    "1": 0,
    "0.5": 0,
    "0.2": 0,
    "0.1": 0,
  },
  maxCashByDenomination: {
    "500": 0,
    "200": 0,
    "50": 0,
    "20": 0,
    "10": 0,
    "5": 0,
    "2": 0,
    "1": 0,
    "0.5": 0,
    "0.2": 0,
    "0.1": 0,
  },
  minimumCashEmailEnabled: false,
  maximumCashEmailEnabled: false,
  maximumCash: 0,
  managerAlertEmails: [],
};

const DEFAULT_BAKERY_SETTINGS = {
  selectedBase: "base-media",
  hostelCapacity: 83,
  emailRecipients: [],
  emailConfig: {
    provider: "resend",
    smtpHost: "smtp.gmail.com",
    smtpPort: 465,
    smtpSecure: true,
    smtpUser: "",
    smtpPassword: "",
    fromEmail: "",
    fromName: "Lisboa Central Hostel",
  },
  breadTable: [],
  breadTypes: [
    { id: "bread-type-1", name: "Carcaças", percentage: 42 },
    { id: "bread-type-2", name: "Pão de Centeio 35gr", percentage: 58 },
  ],
};

const DEFAULT_LAUNDRY_SETTINGS = {
  pricePerKg: 0,
  emailRecipients: [],
  emailEnabled: false,
  emailTime: "00:00",
  managementEmailRecipients: [],
  managementEmailEnabled: false,
  managementEmailTime: "00:00",
  itemTypes: [
    { id: "single-baixo", name: "single baixo", weightKg: 0.48 },
    { id: "single-cima", name: "single cima", weightKg: 0.5 },
    { id: "casal-baixo", name: "casal baixo", weightKg: 0.72 },
    { id: "casal-cima", name: "casal cima", weightKg: 0.75 },
  ],
};

const DEFAULT_HOURS_SETTINGS = {
  people: ["Fernanda Pereira"],
};

const GUESTS_SCREEN_FIELD_OPTIONS = [
  { key: "name", label: "Name" },
  { key: "nationality", label: "Nationality" },
  { key: "birthDate", label: "Birth Date" },
  { key: "docNumber", label: "Doc. Number" },
  { key: "docType", label: "Doc Type" },
  { key: "issuerCountry", label: "Issuer Country" },
  { key: "checkIn", label: "Check-in" },
  { key: "checkOut", label: "Check-out" },
];

const GUESTS_INTEGRATION_MAPPING_ROWS = [
  { key: "name", label: "Name" },
  { key: "nationality", label: "Nationality ICAO" },
  { key: "birthDate", label: "Birth Day" },
  { key: "docNumber", label: "Doc Number" },
  { key: "docType", label: "Doc Type" },
  { key: "issuerCountry", label: "Issuer Country ICAO" },
  { key: "residenceCountry", label: "Residence Country" },
  { key: "residenceCity", label: "Residence City" },
  { key: "checkIn", label: "Check-in date" },
  { key: "checkOut", label: "Check-out date" },
];

const DEFAULT_GUESTS_INTEGRATION_MAPPING = {
  name: "name",
  nationality: "nationality",
  birthDate: "birthDate",
  docNumber: "docNumber",
  docType: "docType",
  issuerCountry: "issuerCountry",
  residenceCountry: "issuerCountry",
  residenceCity: "issuerCountry",
  checkIn: "checkIn",
  checkOut: "checkOut",
};

const DEFAULT_GUESTS_SETTINGS = {
  sendTime: "18:00",
  integrationMapping: { ...DEFAULT_GUESTS_INTEGRATION_MAPPING },
  sefCredentials: {
    unitCode: "508459893",
    establishment: "00",
    accessKey: "102907025181",
    caCertificate: "",
  },
};

const GUEST_DESCRIPTION_PALETTES = {
  blue: { solid: "#d9e1f2", soft: "rgba(217, 225, 242, 0.26)" },
  pink: { solid: "#e890ab", soft: "rgba(232, 144, 171, 0.26)" },
  rose: { solid: "#eeb0c3", soft: "rgba(238, 176, 195, 0.26)" },
  yellow: { solid: "#ffd999", soft: "rgba(255, 217, 153, 0.26)" },
  green: { solid: "#c6deb5", soft: "rgba(198, 222, 181, 0.26)" },
};

const PROFILE_MATRIX_ROWS = [
  { label: "Profile Name", kind: "meta", key: "name" },
  { label: "App: Communications", kind: "app", key: "communications" },
  { label: "App: Guests", kind: "app", key: "guests" },
  { label: "App: Cash Control", kind: "app", key: "cash" },
  { label: "App: Lost&Found", kind: "app", key: "lost-found" },
  { label: "App: Reviews", kind: "app", key: "reviews" },
  { label: "App: Groups", kind: "app", key: "groups" },
  { label: "App: Services", kind: "app", key: "services" },
  { label: "App: Shopping", kind: "app", key: "shopping" },
  { label: "App: Hours Register", kind: "app", key: "hours" },
  { label: "App: Bakery", kind: "app", key: "bakery" },
  { label: "App: Laundry Control", kind: "app", key: "laundry" },
  { label: "Settings: General", kind: "settings", key: "general" },
  { label: "Settings: Communications", kind: "settings", key: "communications" },
  { label: "Settings: Guests", kind: "settings", key: "guests" },
  { label: "Settings: Cash Control", kind: "settings", key: "cash" },
  { label: "Settings: Reviews", kind: "settings", key: "reviews" },
  { label: "Settings: Groups", kind: "settings", key: "groups" },
  { label: "Settings: Services", kind: "settings", key: "services" },
  { label: "Settings: Shopping", kind: "settings", key: "shopping" },
  { label: "Settings: Hours Register", kind: "settings", key: "hours" },
  { label: "Settings: Bakery", kind: "settings", key: "bakery" },
  { label: "Settings: Laundry Control", kind: "settings", key: "laundry" },
  { label: "Settings: Admin Users", kind: "settings", key: "admin-users" },
  { label: "Action", kind: "action", key: "action" },
];

const GROUP_PROPOSAL_TEMPLATES = {
  pt: `Caro/a {{name}},

Obrigado por nos contactar.
Segue abaixo a nossa proposta com base no seu pedido:

Chegada: {{arrival}}
Partida: {{departure}} ({{nights}} noites)

{{room_table}}

Total Alojamento = {{accommodation_total}}

Taxa Municipal {{guests}} hóspedes x {{city_tax_nights}} noites x 4€ = {{city_tax_total}}

Total = {{total}}

O preço inclui: lençóis, cozinha totalmente equipada, receção 24h, internet gratuita (computadores no lobby e Wi-Fi em todo o hostel) e muita informação sobre Lisboa.
Também oferecemos pequeno-almoço gratuito servido diariamente das 08:00 às 11:00, que inclui uma variedade generosa de opções: três tipos de cereais, três tipos de pão, muffins, mini croissants, compota, mel, manteiga, manteiga de amendoim, creme de chocolate, fruta, café, chá, cacau, leite, sumo e as nossas panquecas caseiras!

É necessário um depósito não reembolsável de {{deposit_percentage}}% ({{deposit_value}}) para confirmar a reserva. O valor restante deve ser pago até {{last_payment_days}} dias antes da chegada.

O pagamento pode ser feito por:
Transferência bancária
Cartão de crédito (podemos enviar um link de pagamento seguro)

Dados bancários:
IBAN: PT50 0035 0137 00004852230 14
BIC/SWIFT: CGDIPTPL

Também podemos disponibilizar tours e atividades como Lisbon Walking Tours, aulas de surf e PubCrawls. Por favor informe-nos se precisar de mais informações.

Política de cancelamento: {{deposit_percentage}}% após a confirmação da reserva, 100% se cancelado a menos de {{last_payment_days}} dias antes do check-in.

Por favor note que existe uma taxa municipal de 4 EUR por pessoa, por noite, aplicável a todos os hóspedes com 13 anos ou mais. O valor desta TAXA já está incluído no preço total acima. Está sujeita a um valor máximo de 28 EUR por hóspede.

Por favor informe-nos se precisar de alguma informação adicional.

Esperamos ter notícias suas em breve,`,
  es: `Estimado/a {{name}},

Gracias por contactarnos.
A continuación encontrará nuestra propuesta basada en su solicitud:

Llegada: {{arrival}}
Salida: {{departure}} ({{nights}} noches)

{{room_table}}

Total Alojamiento = {{accommodation_total}}

Tasa Municipal {{guests}} huéspedes x {{city_tax_nights}} noches x 4€ = {{city_tax_total}}

Total = {{total}}

El precio incluye: sábanas, cocina totalmente equipada, recepción 24h, internet gratuito (ordenadores en el lobby y Wi-Fi en todo el hostel), además de mucha información sobre Lisboa.
También ofrecemos desayuno gratuito todos los días de 08:00 a 11:00, que incluye una generosa variedad de opciones: tres tipos de cereales, tres tipos de pan, muffins, mini croissants, mermelada, miel, mantequilla, mantequilla de cacahuete, crema de chocolate, fruta, café, té, cacao, leche, zumo y nuestras tortitas caseras.

Se requiere un depósito no reembolsable del {{deposit_percentage}}% ({{deposit_value}}) para confirmar la reserva. El importe restante debe pagarse hasta {{last_payment_days}} días antes de la llegada.

El pago se puede realizar mediante:
Transferencia bancaria
Tarjeta de crédito (podemos enviar un enlace de pago seguro)

Datos bancarios:
IBAN: PT50 0035 0137 00004852230 14
BIC/SWIFT: CGDIPTPL

También podemos ofrecer tours y actividades como Lisbon Walking Tours, clases de surf y PubCrawls. Por favor, indíquenos si necesita más información.

Política de cancelación: {{deposit_percentage}}% después de la confirmación de la reserva, 100% si se cancela con menos de {{last_payment_days}} días antes del check-in.

Tenga en cuenta que existe una tasa municipal de 4 EUR por persona, por noche, aplicable a todos los huéspedes de 13 años o más. El importe de esta TASA ya está incluido en el precio total indicado arriba. Está sujeto a un importe máximo de 28 EUR por huésped.

Por favor, indíquenos si necesita información adicional.

Esperamos tener noticias suyas pronto,`,
};

const GROUP_CONFIRMATION_TEMPLATES = {
  first: {
    pt: `Caro/a {{name}},

Obrigado pelo seu contacto.
A sua reserva foi confirmada da seguinte forma:

{{confirmation_table}}

Para fazer qualquer alteração a uma reserva existente, por favor contacte-nos.
Por favor informe-nos também sobre a sua hora prevista de chegada.
Por favor note que qualquer cancelamento deve ser comunicado com pelo menos {{last_payment_days}} dias de antecedência (apenas quartos completos são aceites), caso contrário será cobrado o valor total da reserva.

Os nossos dados bancários são:
IBAN: PT50 0035 0137 00004852230 14
BIC SWIFT: CGDIPTPL

Com os melhores cumprimentos,`,
    es: `Estimado/a {{name}},

Gracias por contactarnos.
Su reserva ha sido confirmada de la siguiente forma:

{{confirmation_table}}

Para realizar cualquier cambio en una reserva existente, por favor contáctenos.
Por favor, indíquenos también su hora prevista de llegada.
Tenga en cuenta que cualquier cancelación debe notificarse con al menos {{last_payment_days}} días de antelación (solo se aceptan habitaciones completas), de lo contrario se cobrará el importe total de la reserva.

Nuestros datos bancarios son:
IBAN: PT50 0035 0137 00004852230 14
BIC SWIFT: CGDIPTPL

Saludos cordiales,`,
  },
  final: {
    pt: `Caro/a {{name}},

Obrigado pelo seu pagamento.
A sua reserva encontra-se totalmente paga e confirmada da seguinte forma:

{{confirmation_table}}

Para fazer qualquer alteração a uma reserva existente, por favor contacte-nos.
Por favor informe-nos também sobre a sua hora prevista de chegada.

Com os melhores cumprimentos,`,
    es: `Estimado/a {{name}},

Gracias por su pago.
Su reserva está totalmente pagada y confirmada de la siguiente forma:

{{confirmation_table}}

Para realizar cualquier cambio en una reserva existente, por favor contáctenos.
Por favor, indíquenos también su hora prevista de llegada.

Saludos cordiales,`,
  },
};

const GROUP_ROOM_TYPE_TRANSLATIONS = {
  pt: {
    "11 Bed Dorm Shared Bathroom": "Dormitório de 11 camas com casa de banho partilhada",
    "10 Bed Dorm Shared Bathroom": "Dormitório de 10 camas com casa de banho partilhada",
    "9 Bed Dorm Shared Bathroom": "Dormitório de 9 camas com casa de banho partilhada",
    "8 Bed Dorm Shared Bathroom": "Dormitório de 8 camas com casa de banho partilhada",
    "7 Bed Dorm Shared Bathroom": "Dormitório de 7 camas com casa de banho partilhada",
    "6 Bed Dorm Shared Bathroom": "Dormitório de 6 camas com casa de banho partilhada",
    "5 Bed Dorm Shared Bathroom": "Dormitório de 5 camas com casa de banho partilhada",
    "4 Bed Dorm Shared Bathroom": "Dormitório de 4 camas com casa de banho partilhada",
    "4 Bed Dorm Private Bathroom": "Dormitório de 4 camas com casa de banho privativa",
    "3 Bed Dorm Shared Bathroom": "Dormitório de 3 camas com casa de banho partilhada",
    "3 Bed Dorm Private Bathroom": "Dormitório de 3 camas com casa de banho privativa",
    "2 Bed Dorm Shared Bathroom": "Dormitório de 2 camas com casa de banho partilhada",
    "2 Bed Dorm Private Bathroom": "Dormitório de 2 camas com casa de banho privativa",
    "Twin Private with Private Bathroom": "Quarto twin privado com casa de banho privativa",
    "Twin Private with Shared Bathroom": "Quarto twin privado com casa de banho partilhada",
    "Single Private with Private Bathroom": "Quarto individual privado com casa de banho privativa",
    "Single Private with Shared Bathroom": "Quarto individual privado com casa de banho partilhada",
  },
  es: {
    "11 Bed Dorm Shared Bathroom": "Dormitorio de 11 camas con baño compartido",
    "10 Bed Dorm Shared Bathroom": "Dormitorio de 10 camas con baño compartido",
    "9 Bed Dorm Shared Bathroom": "Dormitorio de 9 camas con baño compartido",
    "8 Bed Dorm Shared Bathroom": "Dormitorio de 8 camas con baño compartido",
    "7 Bed Dorm Shared Bathroom": "Dormitorio de 7 camas con baño compartido",
    "6 Bed Dorm Shared Bathroom": "Dormitorio de 6 camas con baño compartido",
    "5 Bed Dorm Shared Bathroom": "Dormitorio de 5 camas con baño compartido",
    "4 Bed Dorm Shared Bathroom": "Dormitorio de 4 camas con baño compartido",
    "4 Bed Dorm Private Bathroom": "Dormitorio de 4 camas con baño privado",
    "3 Bed Dorm Shared Bathroom": "Dormitorio de 3 camas con baño compartido",
    "3 Bed Dorm Private Bathroom": "Dormitorio de 3 camas con baño privado",
    "2 Bed Dorm Shared Bathroom": "Dormitorio de 2 camas con baño compartido",
    "2 Bed Dorm Private Bathroom": "Dormitorio de 2 camas con baño privado",
    "Twin Private with Private Bathroom": "Habitación twin privada con baño privado",
    "Twin Private with Shared Bathroom": "Habitación twin privada con baño compartido",
    "Single Private with Private Bathroom": "Habitación individual privada con baño privado",
    "Single Private with Shared Bathroom": "Habitación individual privada con baño compartido",
  },
};

const state = {
  entries: [],
  lostFound: [],
  groups: [],
  reviews: [],
  sidebarReviewSummary: null,
  sidebarReviewSummaryLoaded: false,
  reviewProperties: [],
  reviewImportRuns: [],
  reviewStagingRows: [],
  reviewImportRunId: "",
  reviewImportPastedText: "",
  reviewSources: clone(DEFAULT_REVIEW_SOURCES),
  reviewGoogle: { connected: false, connectedAt: "", locations: [], propertyLocations: {}, status: "" },
  reviewScreen: "list",
  reviewSettingsScreen: "import",
  reviewListPage: 1,
  reviewSelectedId: "",
  reviewQa: { prompt: "", answer: "", status: "", loading: false, analyzedCount: 0, totalCount: 0 },
  reviewFilters: { propertyId: "", source: "", search: "", dateFrom: "", dateTo: "", scoreFrom: "", scoreTo: "" },
  groupDraft: emptyGroupDraft(),
  groupSelectedId: "",
  groupEditorTab: "details",
  groupProposalLanguage: "en",
  groupsScreen: "list",
  groupResumeMonthMode: "created",
  groupSort: { key: "dates", dir: "asc" },
  groupSettingsTab: "config",
  groupSettings: clone(DEFAULT_GROUP_SETTINGS),
  groupsShowActive: true,
  groupFilters: { createdFrom: "", createdTo: "", dateFrom: "", dateTo: "", search: "" },
  services: [],
  servicesLoaded: false,
  serviceSettings: clone(DEFAULT_SERVICE_SETTINGS),
  serviceSettingsLoaded: false,
  shoppingOpenOrder: null,
  shoppingHistory: [],
  shoppingLoaded: false,
  shoppingSettings: clone(DEFAULT_SHOPPING_SETTINGS),
  shoppingSettingsLoaded: false,
  shoppingTab: "current",
  shoppingFilters: { category: "", stored: "", groupBy: "category" },
  shoppingHistoryFilters: { dateFrom: "", dateTo: "", name: "", category: "", supplier: "" },
  shoppingSubmitName: "",
  shoppingSubmitNotes: "",
  shoppingSubmitPromptOpen: false,
  shoppingSelectedHistoryId: "",
  cashRecords: [],
  cashLoaded: false,
  cashSettings: clone(DEFAULT_CASH_SETTINGS),
  cashSettingsLoaded: false,
  cashSettingsTab: "config",
  cashScreen: "list",
  cashFilters: { dateFrom: cashDefaultDateFrom(), dateTo: "", shift: "", name: "" },
  cashDraft: null,
  cashOpenDraft: null,
  cashEditDraft: null,
  cashEditingId: "",
  cashMoneyModalOpen: false,
  cashMoneyModalScope: "new",
  cashMoneyModalId: "",
  cashItemsModalOpen: false,
  cashItemsModalScope: "new",
  cashItemsModalId: "",
  cashItemsDraft: {},
  cashItemsJustificationsDraft: {},
  guestsRows: [],
  guestDescriptionRows: [],
  guestsBlacklist: [],
  guestsCountries: [],
  guestsApiCalls: [],
  guestsApiCallsEnabled: true,
  guestsLoaded: false,
  guestsSettings: clone(DEFAULT_GUESTS_SETTINGS),
  guestsSettingsLoaded: false,
  guestsSettingsTab: "config",
  guestsScreen: "list",
  guestsFilters: { showActive: true, ha: "", search: "", nationality: "", checkInFrom: "", checkInTo: "", checkOutFrom: "", checkOutTo: "" },
  guestsDescriptionsFilters: { room: "", description: "" },
  guestsBlacklistFilters: { search: "", whoReported: "", nationality: "" },
  guestsDraft: null,
  guestsEditDraft: null,
  guestsEditingId: "",
  guestsQuickEditId: "",
  guestsQuickEditField: "",
  guestsBlacklistDraft: null,
  guestsBlacklistEditDraft: null,
  guestsBlacklistEditingId: "",
  hoursRecords: [],
  hoursLoaded: false,
  hoursSettings: clone(DEFAULT_HOURS_SETTINGS),
  hoursSettingsLoaded: false,
  hoursScreen: "list",
  hoursFilters: { person: "", dateFrom: "", dateTo: "" },
  hoursDraft: null,
  hoursEditDraft: null,
  hoursEditingId: null,
  bakeryOpenOrder: null,
  bakeryHistory: [],
  bakeryLoaded: false,
  bakerySettings: clone(DEFAULT_BAKERY_SETTINGS),
  bakerySettingsLoaded: false,
  bakeryTab: "current",
  bakerySubmitName: "",
  bakerySelectedHistoryId: "",
  bakerySettingsTab: "table",
  laundryRecords: [],
  laundryLoaded: false,
  laundrySettings: clone(DEFAULT_LAUNDRY_SETTINGS),
  laundrySettingsLoaded: false,
  laundryScreen: "list",
  laundryFilters: { property: "", dateFrom: "", dateTo: "", search: "" },
  laundryResumeFilters: { dateField: "sent", property: "", dateFrom: `${new Date().getFullYear()}-01-01`, dateTo: "", detail: false },
  laundryDraft: null,
  laundrySelectedId: "",
  serviceProviders: [],
  serviceFilters: { showActive: true, createdFrom: "", createdTo: "", dateFrom: "", dateTo: "", name: "" },
  serviceDraft: emptyServiceDraft(),
  serviceSelectedId: "",
  pendingServiceDeepLinkId: "",
  serviceFlightStatuses: {
    cache: {},
    timer: null,
    sequence: 0,
    initialized: false,
  },
  serviceInlineStatusSaving: {},
  servicesScreen: "list",
  serviceDraftFlightPredictions: {
    cache: {},
    timer: null,
    main: { key: "", text: "" },
    return: { key: "", text: "" },
  },
  serviceEditorTab: "details",
  serviceConfirmationLanguage: "en",
  serviceSettingsTab: "config",
  serviceSettingsTemplateType: "",
  serviceSettingsTemplateLanguage: "en",
  editingId: null,
  newDraft: { person: "", status: "Open", category: "Information", message: "" },
  editDraft: null,
  lostFoundEditingId: null,
  lostFoundDraft: emptyLostFoundDraft(),
  lostFoundEditDraft: null,
  sort: { key: "date", dir: "desc" },
  pendingDelete: null,
  access: {
    profile: { id: "", name: "Full access" },
    appFeatures: [...APP_FEATURE_OPTIONS],
    settingsFeatures: [...SETTINGS_FEATURE_OPTIONS],
  },
  profiles: [],
  profilesLoaded: false,
  settings: clone(DEFAULT_SETTINGS),
  currentView: "communications",
  settingsSection: "general",
  autoRefreshTimer: null,
  lastAutoRefreshAt: 0,
  autoRefreshRunning: false,
  mobileNavOpen: false,
  adminUsers: [],
  adminUsersLoaded: false,
  communicationsLoaded: false,
  lostFoundLoaded: false,
  communicationsSettingsLoaded: false,
  groupsLoaded: false,
  groupSettingsLoaded: false,
  reviewDateFilterApplied: false,
  reviewPropertiesLoaded: false,
  reviewSettingsLoaded: false,
  reviewGoogleLoaded: false,
  reviewsLoaded: false,
  reviewImportRunsLoaded: false,
  supabase: null,
  user: null,
};

const els = {
  appShell: document.getElementById("app-shell"),
  leftNav: document.querySelector(".left-nav"),
  sidebarReviewSummaryCard: document.getElementById("sidebar-review-summary-card"),
  sidebarReviewSummaryStatus: document.getElementById("sidebar-review-summary-status"),
  sidebarReviewSummaryBody: document.getElementById("sidebar-review-summary-body"),
  topbar: document.querySelector(".topbar"),
  mobileMenuToggle: document.getElementById("mobile-menu-toggle"),
  navCommunications: document.getElementById("nav-communications"),
  navGuests: document.getElementById("nav-guests"),
  navCash: document.getElementById("nav-cash"),
  navLostFound: document.getElementById("nav-lost-found"),
  navReviews: document.getElementById("nav-reviews"),
  navGroups: document.getElementById("nav-groups"),
  navServices: document.getElementById("nav-services"),
  navShopping: document.getElementById("nav-shopping"),
  navHours: document.getElementById("nav-hours"),
  navBakery: document.getElementById("nav-bakery"),
  navLaundry: document.getElementById("nav-laundry"),
  openSettings: document.getElementById("open-settings"),
  closeSettings: document.getElementById("close-settings"),
  viewCommunications: document.getElementById("view-communications"),
  viewGuests: document.getElementById("view-guests"),
  viewCash: document.getElementById("view-cash"),
  viewLostFound: document.getElementById("view-lost-found"),
  viewReviews: document.getElementById("view-reviews"),
  viewServices: document.getElementById("view-services"),
  viewShopping: document.getElementById("view-shopping"),
  viewHours: document.getElementById("view-hours"),
  viewBakery: document.getElementById("view-bakery"),
  viewLaundry: document.getElementById("view-laundry"),
  viewSettings: document.getElementById("view-settings"),
  settingsMenuGeneral: document.getElementById("settings-menu-general"),
  settingsMenuCommunications: document.getElementById("settings-menu-communications"),
  settingsMenuGuests: document.getElementById("settings-menu-guests"),
  settingsMenuCash: document.getElementById("settings-menu-cash"),
  settingsMenuReviews: document.getElementById("settings-menu-reviews"),
  settingsMenuGroups: document.getElementById("settings-menu-groups"),
  settingsMenuServices: document.getElementById("settings-menu-services"),
  settingsMenuShopping: document.getElementById("settings-menu-shopping"),
  settingsMenuHours: document.getElementById("settings-menu-hours"),
  settingsMenuBakery: document.getElementById("settings-menu-bakery"),
  settingsMenuLaundry: document.getElementById("settings-menu-laundry"),
  settingsMenuAdminUsers: document.getElementById("settings-menu-admin-users"),
  closeSettingsCash: document.getElementById("close-settings-cash"),
  cashSettingsConfigTab: document.getElementById("cash-settings-config-tab"),
  cashSettingsMinTab: document.getElementById("cash-settings-min-tab"),
  cashSaveSettings: document.getElementById("cash-save-settings"),
  cashAddShift: document.getElementById("cash-add-shift"),
  cashAddItem: document.getElementById("cash-add-item"),
  cashSettingsConfigPanel: document.getElementById("cash-settings-config-panel"),
  cashSettingsConfigItemsPanel: document.getElementById("cash-settings-config-items-panel"),
  cashSettingsMinPanel: document.getElementById("cash-settings-min-panel"),
  cashSettingsShiftsBody: document.getElementById("cash-settings-shifts-body"),
  cashSettingsItemsBody: document.getElementById("cash-settings-items-body"),
  cashSettingsMinBody: document.getElementById("cash-settings-min-body"),
  cashSettingsManagerAlertEmail: document.getElementById("cash-settings-manager-alert-email"),
  cashSettingsMinimumEmailEnabled: document.getElementById("cash-settings-minimum-email-enabled"),
  cashSettingsMaximumEmailEnabled: document.getElementById("cash-settings-maximum-email-enabled"),
  cashSettingsMaximumCash: document.getElementById("cash-settings-maximum-cash"),
  cashSettingsStatus: document.getElementById("cash-settings-status"),
  cashTabList: document.getElementById("cash-tab-list"),
  cashTabDetail: document.getElementById("cash-tab-detail"),
  cashTabItems: document.getElementById("cash-tab-items"),
  cashTabResume: document.getElementById("cash-tab-resume"),
  cashCount: document.getElementById("cash-count"),
  cashWarning: document.getElementById("cash-warning"),
  cashFilterDateFrom: document.getElementById("cash-filter-date-from"),
  cashFilterDateTo: document.getElementById("cash-filter-date-to"),
  cashFilterShift: document.getElementById("cash-filter-shift"),
  cashFilterName: document.getElementById("cash-filter-name"),
  cashPanelList: document.getElementById("cash-panel-list"),
  cashPanelDetail: document.getElementById("cash-panel-detail"),
  cashPanelItems: document.getElementById("cash-panel-items"),
  cashPanelResume: document.getElementById("cash-panel-resume"),
  cashRows: document.getElementById("cash-rows"),
  cashDetailRows: document.getElementById("cash-detail-rows"),
  cashItemDetailHead: document.getElementById("cash-item-detail-head"),
  cashItemDetailRows: document.getElementById("cash-item-detail-rows"),
  cashResumeRows: document.getElementById("cash-resume-rows"),
  cashMobileCards: document.getElementById("cash-mobile-cards"),
  cashStatus: document.getElementById("cash-status"),
  cashMoneyModal: document.getElementById("cash-money-modal"),
  cashMoneyClose: document.getElementById("cash-money-close"),
  cashMoneyMeta: document.getElementById("cash-money-meta"),
  cashMoneyBody: document.getElementById("cash-money-body"),
  cashMoneyTotal: document.getElementById("cash-money-total"),
  cashMoneySave: document.getElementById("cash-money-save"),
  cashItemsModal: document.getElementById("cash-items-modal"),
  cashItemsClose: document.getElementById("cash-items-close"),
  cashItemsMeta: document.getElementById("cash-items-meta"),
  cashItemsBody: document.getElementById("cash-items-body"),
  cashItemsStatus: document.getElementById("cash-items-status"),
  cashItemsSave: document.getElementById("cash-items-save"),
  settingsViewGeneral: document.getElementById("settings-view-general"),
  settingsViewCommunications: document.getElementById("settings-view-communications"),
  settingsViewGuests: document.getElementById("settings-view-guests"),
  settingsViewCash: document.getElementById("settings-view-cash"),
  settingsViewReviews: document.getElementById("settings-view-reviews"),
  settingsViewGroups: document.getElementById("settings-view-groups"),
  settingsViewServices: document.getElementById("settings-view-services"),
  settingsViewShopping: document.getElementById("settings-view-shopping"),
  settingsViewHours: document.getElementById("settings-view-hours"),
  settingsViewBakery: document.getElementById("settings-view-bakery"),
  settingsViewLaundry: document.getElementById("settings-view-laundry"),
  settingsViewAdminUsers: document.getElementById("settings-view-admin-users"),
  settingsReviewsImportTab: document.getElementById("settings-reviews-import-tab"),
  settingsReviewsConfigTab: document.getElementById("settings-reviews-config-tab"),
  settingsReviewsImportPanel: document.getElementById("settings-reviews-import-panel"),
  settingsReviewsConfigPanel: document.getElementById("settings-reviews-config-panel"),
  closeSettingsGeneral: document.getElementById("close-settings-general"),
  closeSettingsGuests: document.getElementById("close-settings-guests"),
  closeSettingsAdmin: document.getElementById("close-settings-admin"),
  closeSettingsReviews: document.getElementById("close-settings-reviews"),
  closeSettingsGroups: document.getElementById("close-settings-groups"),
  closeSettingsServices: document.getElementById("close-settings-services"),
  closeSettingsShopping: document.getElementById("close-settings-shopping"),
  closeSettingsHours: document.getElementById("close-settings-hours"),
  closeSettingsBakery: document.getElementById("close-settings-bakery"),
  closeSettingsLaundry: document.getElementById("close-settings-laundry"),
  generalSaveSettings: document.getElementById("general-save-settings"),
  generalEmailProvider: document.getElementById("general-email-provider"),
  generalEmailSmtpHost: document.getElementById("general-email-smtp-host"),
  generalEmailSmtpPort: document.getElementById("general-email-smtp-port"),
  generalEmailSmtpSecure: document.getElementById("general-email-smtp-secure"),
  generalEmailSmtpUser: document.getElementById("general-email-smtp-user"),
  generalEmailSmtpPassword: document.getElementById("general-email-smtp-password"),
  generalEmailFromEmail: document.getElementById("general-email-from-email"),
  generalEmailFromName: document.getElementById("general-email-from-name"),
  generalEmailSmtpFields: document.getElementById("general-email-smtp-fields"),
  generalSettingsStatus: document.getElementById("general-settings-status"),
  adminUserEmail: document.getElementById("admin-user-email"),
  adminUserPassword: document.getElementById("admin-user-password"),
  adminUserProfile: document.getElementById("admin-user-profile"),
  adminCreateUser: document.getElementById("admin-create-user"),
  adminRefreshUsers: document.getElementById("admin-refresh-users"),
  adminUsersStatus: document.getElementById("admin-users-status"),
  adminUsersBody: document.getElementById("admin-users-body"),
  profilesHead: document.getElementById("profiles-head"),
  profilesBody: document.getElementById("profiles-body"),
  addProfile: document.getElementById("add-profile"),
  profilesStatus: document.getElementById("profiles-status"),
  rows: document.getElementById("rows"),
  communicationsMobileCards: document.getElementById("communications-mobile-cards"),
  lostFoundRows: document.getElementById("lost-found-rows"),
  lostFoundMobileCards: document.getElementById("lost-found-mobile-cards"),
  lostFoundCount: document.getElementById("lost-found-count"),
  lostFoundDbStatus: document.getElementById("lost-found-status"),
  lostFoundOnlyOpen: document.getElementById("lost-found-only-open"),
  lostFoundFilterNumber: document.getElementById("lost-found-filter-number"),
  lostFoundFilterDate: document.getElementById("lost-found-filter-date"),
  lostFoundFilterWhoFound: document.getElementById("lost-found-filter-who-found"),
  lostFoundFilterWhoRecorded: document.getElementById("lost-found-filter-who-recorded"),
  lostFoundFilterWhere: document.getElementById("lost-found-filter-where"),
  lostFoundFilterObject: document.getElementById("lost-found-filter-object"),
  lostFoundFilterNotes: document.getElementById("lost-found-filter-notes"),
  lostFoundFilterStored: document.getElementById("lost-found-filter-stored"),
  tableWrap: document.getElementById("communications-table-wrap"),
  tableHead: document.getElementById("communications-head"),
  resetSort: document.getElementById("reset-sort"),
  count: document.getElementById("count"),
  search: document.getElementById("communications-search"),
  showActive: document.getElementById("show-active"),
  groupCommunications: document.getElementById("communications-group"),
  statusFilter: document.getElementById("status-filter"),
  categoryFilter: document.getElementById("category-filter"),
  fromDate: document.getElementById("from-date"),
  toDate: document.getElementById("to-date"),
  excelInput: document.getElementById("excel-input"),
  exportCsv: document.getElementById("export-csv"),
  dbStatus: document.getElementById("db-status"),
  authLogout: document.getElementById("auth-logout"),
  authUser: document.getElementById("auth-user"),
  settingsCategoriesBody: document.getElementById("settings-categories-body"),
  addCategory: document.getElementById("add-category"),
  settingEmailEnabled: document.getElementById("setting-email-enabled"),
  settingEmailFrequency: document.getElementById("setting-email-frequency"),
  settingEmailTime: document.getElementById("setting-email-time"),
  settingEmailRecipients: document.getElementById("setting-email-recipients"),
  settingEmailFrequency2: document.getElementById("setting-email-frequency-2"),
  settingEmailTime2: document.getElementById("setting-email-time-2"),
  settingEmailRecipients2: document.getElementById("setting-email-recipients-2"),
  settingEmailPreview: document.getElementById("setting-email-preview"),
  settingEmailNextPreview: document.getElementById("setting-email-next-preview"),
  settingEmailTestRecipient: document.getElementById("setting-email-test-recipient"),
  testEmailNow: document.getElementById("test-email-now"),
  saveSettings: document.getElementById("save-settings"),
  settingsStatus: document.getElementById("settings-status"),
  guestsTabList: document.getElementById("guests-tab-list"),
  guestsTabDescriptions: document.getElementById("guests-tab-descriptions"),
  guestsTabBlacklist: document.getElementById("guests-tab-blacklist"),
  guestsPanelList: document.getElementById("guests-panel-list"),
  guestsPanelDescriptions: document.getElementById("guests-panel-descriptions"),
  guestsPanelBlacklist: document.getElementById("guests-panel-blacklist"),
  guestsListAlertReason: document.getElementById("guests-list-alert-reason"),
  guestsShowActive: document.getElementById("guests-show-active"),
  guestsExportExcel: document.getElementById("guests-export-excel"),
  guestsSendPending: document.getElementById("guests-send-pending"),
  guestsCount: document.getElementById("guests-count"),
  guestsFilterHa: document.getElementById("guests-filter-ha"),
  guestsFilterSearch: document.getElementById("guests-filter-search"),
  guestsFilterNationality: document.getElementById("guests-filter-nationality"),
  guestsFilterCheckinFrom: document.getElementById("guests-filter-checkin-from"),
  guestsFilterCheckinTo: document.getElementById("guests-filter-checkin-to"),
  guestsFilterCheckoutFrom: document.getElementById("guests-filter-checkout-from"),
  guestsFilterCheckoutTo: document.getElementById("guests-filter-checkout-to"),
  guestsAlertSummary: document.getElementById("guests-alert-summary"),
  guestsCountryList: document.getElementById("guests-country-list"),
  guestsRows: document.getElementById("guests-rows"),
  guestsDescriptionsRows: document.getElementById("guests-descriptions-rows"),
  guestsMobileCards: document.getElementById("guests-mobile-cards"),
  guestsDescriptionsMobileCards: document.getElementById("guests-descriptions-mobile-cards"),
  guestsStatus: document.getElementById("guests-status"),
  guestsDescriptionsCount: document.getElementById("guests-descriptions-count"),
  guestsDescriptionsStatus: document.getElementById("guests-descriptions-status"),
  guestsDescriptionsFilterRoom: document.getElementById("guests-descriptions-filter-room"),
  guestsDescriptionsFilterDescription: document.getElementById("guests-descriptions-filter-description"),
  guestsBlacklistCount: document.getElementById("guests-blacklist-count"),
  guestsBlacklistFilterSearch: document.getElementById("guests-blacklist-filter-search"),
  guestsBlacklistFilterReported: document.getElementById("guests-blacklist-filter-reported"),
  guestsBlacklistFilterNationality: document.getElementById("guests-blacklist-filter-nationality"),
  guestsBlacklistRows: document.getElementById("guests-blacklist-rows"),
  guestsBlacklistMobileCards: document.getElementById("guests-blacklist-mobile-cards"),
  guestsBlacklistStatus: document.getElementById("guests-blacklist-status"),
  guestsSaveSettings: document.getElementById("guests-save-settings"),
  guestsSettingsConfigTab: document.getElementById("guests-settings-config-tab"),
  guestsSettingsSefTab: document.getElementById("guests-settings-sef-tab"),
  guestsSettingsApiTab: document.getElementById("guests-settings-api-tab"),
  guestsSettingsConfigPanel: document.getElementById("guests-settings-config-panel"),
  guestsSettingsSefPanel: document.getElementById("guests-settings-sef-panel"),
  guestsSettingsApiPanel: document.getElementById("guests-settings-api-panel"),
  guestsSettingsSendTime: document.getElementById("guests-settings-send-time"),
  guestsSettingsMappingBody: document.getElementById("guests-settings-mapping-body"),
  guestsSettingsSefUnit: document.getElementById("guests-settings-sef-unit"),
  guestsSettingsSefEstablishment: document.getElementById("guests-settings-sef-establishment"),
  guestsSettingsSefAccessKey: document.getElementById("guests-settings-sef-access-key"),
  guestsSettingsSefCa: document.getElementById("guests-settings-sef-ca"),
  guestsSettingsApiNote: document.getElementById("guests-settings-api-note"),
  guestsSettingsApiBody: document.getElementById("guests-settings-api-body"),
  guestsSettingsStatus: document.getElementById("guests-settings-status"),
  viewGroups: document.getElementById("view-groups"),
  groupsNew: document.getElementById("groups-new"),
  groupsTabList: document.getElementById("groups-tab-list"),
  groupsTabResume: document.getElementById("groups-tab-resume"),
  groupsStatus: document.getElementById("groups-status"),
  groupReservationNumber: document.getElementById("group-reservation-number"),
  groupName: document.getElementById("group-name"),
  groupEmail: document.getElementById("group-email"),
  groupEmailProposalsHint: document.getElementById("group-email-proposals-hint"),
  groupCheckIn: document.getElementById("group-check-in"),
  groupCheckOut: document.getElementById("group-check-out"),
  groupNightsLabel: document.getElementById("group-nights-label"),
  groupGuests: document.getElementById("group-guests"),
  groupOptionDate: document.getElementById("group-option-date"),
  groupStatusField: document.getElementById("group-status-field"),
  groupLastPaymentLimit: document.getElementById("group-last-payment-limit"),
  groupObservation: document.getElementById("group-observation"),
  groupEditorModal: document.getElementById("group-editor-modal"),
  groupTabDetails: document.getElementById("group-tab-details"),
  groupTabEmail: document.getElementById("group-tab-email"),
  groupTabConfirmation: document.getElementById("group-tab-confirmation"),
  groupTabFinalConfirmation: document.getElementById("group-tab-final-confirmation"),
  groupDetailsPanel: document.getElementById("group-details-panel"),
  groupEmailPanel: document.getElementById("group-email-panel"),
  groupEmailTitle: document.getElementById("group-email-title"),
  groupEmailDescription: document.getElementById("group-email-description"),
  groupEmailPreview: document.getElementById("group-email-preview"),
  groupProposalLanguage: document.getElementById("group-proposal-language"),
  groupCopyEmail: document.getElementById("group-copy-email"),
  groupCloseModal: document.getElementById("group-close-modal"),
  groupAuditHistory: document.getElementById("group-audit-history"),
  groupRoomItemsBody: document.getElementById("group-room-items-body"),
  groupAddRoomItem: document.getElementById("group-add-room-item"),
  groupGuestCounter: document.getElementById("group-guest-counter"),
  groupAccommodationTotal: document.getElementById("group-accommodation-total"),
  groupCityTaxTotal: document.getElementById("group-city-tax-total"),
  groupTotalValue: document.getElementById("group-total-value"),
  groupDepositPreview: document.getElementById("group-deposit-preview"),
  groupSave: document.getElementById("group-save"),
  groupDelete: document.getElementById("group-delete"),
  groupsShowActive: document.getElementById("groups-show-active"),
  groupsExportExcel: document.getElementById("groups-export-excel"),
  groupsExportPdf: document.getElementById("groups-export-pdf"),
  groupsCount: document.getElementById("groups-count"),
  groupsPanelList: document.getElementById("groups-panel-list"),
  groupsPanelResume: document.getElementById("groups-panel-resume"),
  groupsResumeMonthMode: document.getElementById("groups-resume-month-mode"),
  groupsResumeCount: document.getElementById("groups-resume-count"),
  groupsResumeBody: document.getElementById("groups-resume-body"),
  groupsRows: document.getElementById("groups-rows"),
  groupsMobileCards: document.getElementById("groups-mobile-cards"),
  groupsStatusFooter: document.getElementById("groups-status-footer"),
  groupsFilterCreatedFrom: document.getElementById("groups-filter-created-from"),
  groupsFilterCreatedTo: document.getElementById("groups-filter-created-to"),
  groupsFilterDateFrom: document.getElementById("groups-filter-date-from"),
  groupsFilterDateTo: document.getElementById("groups-filter-date-to"),
  groupsFilterSearch: document.getElementById("groups-filter-search"),
  groupsDepositPercentage: document.getElementById("groups-deposit-percentage"),
  groupsLastPaymentDays: document.getElementById("groups-last-payment-days"),
  groupsEmailTemplate: document.getElementById("groups-email-template"),
  groupsConfirmationTemplate: document.getElementById("groups-confirmation-template"),
  groupsFinalConfirmationTemplate: document.getElementById("groups-final-confirmation-template"),
  groupsProposalTemplatePreview: document.getElementById("groups-proposal-template-preview"),
  groupsConfirmationTemplatePreview: document.getElementById("groups-confirmation-template-preview"),
  groupsFinalConfirmationTemplatePreview: document.getElementById("groups-final-confirmation-template-preview"),
  groupsSettingsConfigTab: document.getElementById("groups-settings-config-tab"),
  groupsSettingsProposalTab: document.getElementById("groups-settings-proposal-tab"),
  groupsSettingsConfirmationTab: document.getElementById("groups-settings-confirmation-tab"),
  groupsSettingsFinalConfirmationTab: document.getElementById("groups-settings-final-confirmation-tab"),
  groupsSettingsConfigPanel: document.getElementById("groups-settings-config-panel"),
  groupsSettingsProposalPanel: document.getElementById("groups-settings-proposal-panel"),
  groupsSettingsConfirmationPanel: document.getElementById("groups-settings-confirmation-panel"),
  groupsSettingsFinalConfirmationPanel: document.getElementById("groups-settings-final-confirmation-panel"),
  groupsRoomTypesBody: document.getElementById("groups-room-types-body"),
  groupsAddRoomType: document.getElementById("groups-add-room-type"),
  groupsSaveSettings: document.getElementById("groups-save-settings"),
  groupsSaveSettingsProposal: document.getElementById("groups-save-settings-proposal"),
  groupsSaveSettingsConfirmation: document.getElementById("groups-save-settings-confirmation"),
  groupsSaveSettingsFinalConfirmation: document.getElementById("groups-save-settings-final-confirmation"),
  groupsSettingsStatus: document.getElementById("groups-settings-status"),
  servicesNew: document.getElementById("services-new"),
  servicesTabList: document.getElementById("services-tab-list"),
  servicesTabResume: document.getElementById("services-tab-resume"),
  servicesPanelList: document.getElementById("services-panel-list"),
  servicesPanelResume: document.getElementById("services-panel-resume"),
  servicesRows: document.getElementById("services-rows"),
  servicesMobileCards: document.getElementById("services-mobile-cards"),
  servicesCount: document.getElementById("services-count"),
  servicesResumeCount: document.getElementById("services-resume-count"),
  servicesResumeBody: document.getElementById("services-resume-body"),
  servicesShowActive: document.getElementById("services-show-active"),
  servicesFilterCreatedFrom: document.getElementById("services-filter-created-from"),
  servicesFilterCreatedTo: document.getElementById("services-filter-created-to"),
  servicesFilterDateFrom: document.getElementById("services-filter-date-from"),
  servicesFilterDateTo: document.getElementById("services-filter-date-to"),
  servicesFilterName: document.getElementById("services-filter-name"),
  servicesDbStatus: document.getElementById("services-db-status"),
  serviceEditorModal: document.getElementById("service-editor-modal"),
  serviceTabDetails: document.getElementById("service-tab-details"),
  serviceTabConfirmation: document.getElementById("service-tab-confirmation"),
  serviceDetailsPanel: document.getElementById("service-details-panel"),
  serviceConfirmationPanel: document.getElementById("service-confirmation-panel"),
  serviceConfirmationLanguage: document.getElementById("service-confirmation-language"),
  serviceCopyConfirmation: document.getElementById("service-copy-confirmation"),
  serviceConfirmationPreview: document.getElementById("service-confirmation-preview"),
  serviceCloseModal: document.getElementById("service-close-modal"),
  serviceRequestNumberLabel: document.getElementById("service-request-number-label"),
  serviceType: document.getElementById("service-type"),
  serviceStatus: document.getElementById("service-status"),
  serviceCustomerName: document.getElementById("service-customer-name"),
  serviceCustomerEmail: document.getElementById("service-customer-email"),
  serviceCustomerPhoneFlag: document.getElementById("service-customer-phone-flag"),
  serviceCustomerPhone: document.getElementById("service-customer-phone"),
  servicePax: document.getElementById("service-pax"),
  serviceDate: document.getElementById("service-date"),
  serviceDatePicker: document.getElementById("service-date-picker"),
  serviceTime: document.getElementById("service-time"),
  serviceTimePrediction: document.getElementById("service-time-prediction"),
  servicePickupLocation: document.getElementById("service-pickup-location"),
  serviceDropoffLocation: document.getElementById("service-dropoff-location"),
  serviceFlightField: document.getElementById("service-flight-field"),
  serviceFlightNumber: document.getElementById("service-flight-number"),
  serviceHasReturn: document.getElementById("service-has-return"),
  servicePrice: document.getElementById("service-price"),
  serviceProviderEmail: document.getElementById("service-provider-email"),
  serviceNotes: document.getElementById("service-notes"),
  serviceReturnFields: document.getElementById("service-return-fields"),
  serviceReturnPickup: document.getElementById("service-return-pickup"),
  serviceReturnDropoff: document.getElementById("service-return-dropoff"),
  serviceReturnDate: document.getElementById("service-return-date"),
  serviceReturnDatePicker: document.getElementById("service-return-date-picker"),
  serviceReturnTime: document.getElementById("service-return-time"),
  serviceReturnTimePrediction: document.getElementById("service-return-time-prediction"),
  serviceReturnFlightField: document.getElementById("service-return-flight-field"),
  serviceReturnFlight: document.getElementById("service-return-flight"),
  serviceAuditHistory: document.getElementById("service-audit-history"),
  serviceSave: document.getElementById("service-save"),
  serviceDelete: document.getElementById("service-delete"),
  servicesStatus: document.getElementById("services-status"),
  servicesExportExcel: document.getElementById("services-export-excel"),
  servicesSaveSettings: document.getElementById("services-save-settings"),
  servicesSaveSettingsConfirmation: document.getElementById("services-save-settings-confirmation"),
  servicesSettingsConfigTab: document.getElementById("services-settings-config-tab"),
  servicesSettingsConfirmationTab: document.getElementById("services-settings-confirmation-tab"),
  servicesSettingsConfigPanel: document.getElementById("services-settings-config-panel"),
  servicesSettingsConfirmationPanel: document.getElementById("services-settings-confirmation-panel"),
  servicesConfigsBody: document.getElementById("services-configs-body"),
  servicesAutomaticEmailRecipients: document.getElementById("services-automatic-email-recipients"),
  servicesLiveFlightStatusEnabled: document.getElementById("services-live-flight-status-enabled"),
  servicesTemplateServiceType: document.getElementById("services-template-service-type"),
  servicesTemplateLanguage: document.getElementById("services-template-language"),
  servicesConfirmationTemplate: document.getElementById("services-confirmation-template"),
  servicesConfirmationTemplatePreview: document.getElementById("services-confirmation-template-preview"),
  servicesSettingsStatus: document.getElementById("services-settings-status"),
  servicesPriceOneWay13: document.getElementById("services-price-oneway-1-3"),
  servicesPriceOneWay47: document.getElementById("services-price-oneway-4-7"),
  servicesPriceOneWay811: document.getElementById("services-price-oneway-8-11"),
  servicesPriceOneWay1216: document.getElementById("services-price-oneway-12-16"),
  servicesPriceReturn13: document.getElementById("services-price-return-1-3"),
  servicesPriceReturn47: document.getElementById("services-price-return-4-7"),
  servicesPriceReturn811: document.getElementById("services-price-return-8-11"),
  servicesPriceReturn1216: document.getElementById("services-price-return-12-16"),
  shoppingTabCurrent: document.getElementById("shopping-tab-current"),
  shoppingTabHistory: document.getElementById("shopping-tab-history"),
  shoppingPanelCurrent: document.getElementById("shopping-panel-current"),
  shoppingPanelHistory: document.getElementById("shopping-panel-history"),
  shoppingNewOrder: document.getElementById("shopping-new-order"),
  shoppingSaveOrder: document.getElementById("shopping-save-order"),
  shoppingOpenSummary: document.getElementById("shopping-open-summary"),
  shoppingCurrentStatus: document.getElementById("shopping-current-status"),
  shoppingOpenEmpty: document.getElementById("shopping-open-empty"),
  shoppingOpenContent: document.getElementById("shopping-open-content"),
  shoppingFilterCategory: document.getElementById("shopping-filter-category"),
  shoppingFilterStored: document.getElementById("shopping-filter-stored"),
  shoppingGroupBy: document.getElementById("shopping-group-by"),
  shoppingOpenRows: document.getElementById("shopping-open-rows"),
  shoppingMobileCards: document.getElementById("shopping-mobile-cards"),
  shoppingSubmitNameWrap: document.getElementById("shopping-submit-name-wrap"),
  shoppingSubmitName: document.getElementById("shopping-submit-name"),
  shoppingSubmitNotes: document.getElementById("shopping-submit-notes"),
  shoppingSubmitStatus: document.getElementById("shopping-submit-status"),
  shoppingSubmitOrder: document.getElementById("shopping-submit-order"),
  shoppingHistoryRows: document.getElementById("shopping-history-rows"),
  shoppingHistoryMobileCards: document.getElementById("shopping-history-mobile-cards"),
  shoppingHistoryCount: document.getElementById("shopping-history-count"),
  shoppingHistoryStatus: document.getElementById("shopping-history-status"),
  shoppingHistoryDateFrom: document.getElementById("shopping-history-date-from"),
  shoppingHistoryDateTo: document.getElementById("shopping-history-date-to"),
  shoppingHistoryName: document.getElementById("shopping-history-name"),
  shoppingHistoryCategory: document.getElementById("shopping-history-category"),
  shoppingHistorySupplier: document.getElementById("shopping-history-supplier"),
  shoppingDetailModal: document.getElementById("shopping-detail-modal"),
  shoppingExportExcel: document.getElementById("shopping-export-excel"),
  shoppingExportPdf: document.getElementById("shopping-export-pdf"),
  shoppingCopyOrder: document.getElementById("shopping-copy-order"),
  shoppingDetailClose: document.getElementById("shopping-detail-close"),
  shoppingReopenOrder: document.getElementById("shopping-reopen-order"),
  shoppingDetailStatus: document.getElementById("shopping-detail-status"),
  shoppingDetailBody: document.getElementById("shopping-detail-body"),
  shoppingSaveSettings: document.getElementById("shopping-save-settings"),
  shoppingSettingsEmailRecipients: document.getElementById("shopping-settings-email-recipients"),
  shoppingSettingsWeekdays: document.getElementById("shopping-settings-weekdays"),
  shoppingSettingsCategoryColors: document.getElementById("shopping-settings-category-colors"),
  shoppingAddItem: document.getElementById("shopping-add-item"),
  shoppingSettingsItemsBody: document.getElementById("shopping-settings-items-body"),
  shoppingSettingsStatus: document.getElementById("shopping-settings-status"),
  hoursExportExcel: document.getElementById("hours-export-excel"),
  hoursTabList: document.getElementById("hours-tab-list"),
  hoursTabResume: document.getElementById("hours-tab-resume"),
  hoursPanelList: document.getElementById("hours-panel-list"),
  hoursPanelResume: document.getElementById("hours-panel-resume"),
  hoursRows: document.getElementById("hours-rows"),
  hoursMobileCards: document.getElementById("hours-mobile-cards"),
  hoursCount: document.getElementById("hours-count"),
  hoursStatus: document.getElementById("hours-status"),
  hoursFilterPerson: document.getElementById("hours-filter-person"),
  hoursFilterDateFrom: document.getElementById("hours-filter-date-from"),
  hoursFilterDateTo: document.getElementById("hours-filter-date-to"),
  hoursResumeFilterPerson: document.getElementById("hours-resume-filter-person"),
  hoursResumeFilterDateFrom: document.getElementById("hours-resume-filter-date-from"),
  hoursResumeFilterDateTo: document.getElementById("hours-resume-filter-date-to"),
  hoursResumeCount: document.getElementById("hours-resume-count"),
  hoursResumeBody: document.getElementById("hours-resume-body"),
  hoursSaveSettings: document.getElementById("hours-save-settings"),
  hoursSettingsPersons: document.getElementById("hours-settings-persons"),
  hoursSettingsStatus: document.getElementById("hours-settings-status"),
  bakeryTabCurrent: document.getElementById("bakery-tab-current"),
  bakeryTabHistory: document.getElementById("bakery-tab-history"),
  bakeryTabResume: document.getElementById("bakery-tab-resume"),
  bakeryPanelCurrent: document.getElementById("bakery-panel-current"),
  bakeryPanelHistory: document.getElementById("bakery-panel-history"),
  bakeryPanelResume: document.getElementById("bakery-panel-resume"),
  bakeryNewOrder: document.getElementById("bakery-new-order"),
  bakerySaveOrder: document.getElementById("bakery-save-order"),
  bakeryOpenSummary: document.getElementById("bakery-open-summary"),
  bakeryCurrentStatus: document.getElementById("bakery-current-status"),
  bakeryOpenEmpty: document.getElementById("bakery-open-empty"),
  bakeryOpenContent: document.getElementById("bakery-open-content"),
  bakeryOpenRows: document.getElementById("bakery-open-rows"),
  bakeryMobileCards: document.getElementById("bakery-mobile-cards"),
  bakeryGeneratedText: document.getElementById("bakery-generated-text"),
  bakerySubmitName: document.getElementById("bakery-submit-name"),
  bakerySubmitStatus: document.getElementById("bakery-submit-status"),
  bakerySubmitOrder: document.getElementById("bakery-submit-order"),
  bakeryHistoryRows: document.getElementById("bakery-history-rows"),
  bakeryHistoryMobileCards: document.getElementById("bakery-history-mobile-cards"),
  bakeryHistoryCount: document.getElementById("bakery-history-count"),
  bakeryHistoryStatus: document.getElementById("bakery-history-status"),
  bakeryResumeHead: document.getElementById("bakery-resume-head"),
  bakeryResumeRows: document.getElementById("bakery-resume-rows"),
  bakeryResumeCount: document.getElementById("bakery-resume-count"),
  bakeryDetailModal: document.getElementById("bakery-detail-modal"),
  bakeryDetailClose: document.getElementById("bakery-detail-close"),
  bakeryDetailResend: document.getElementById("bakery-detail-resend"),
  bakeryDetailStatus: document.getElementById("bakery-detail-status"),
  bakeryDetailBody: document.getElementById("bakery-detail-body"),
  laundryNew: document.getElementById("laundry-new"),
  laundryTabList: document.getElementById("laundry-tab-list"),
  laundryTabResume: document.getElementById("laundry-tab-resume"),
  laundryTabAnalysis: document.getElementById("laundry-tab-analysis"),
  laundryPanelList: document.getElementById("laundry-panel-list"),
  laundryPanelResume: document.getElementById("laundry-panel-resume"),
  laundryPanelAnalysis: document.getElementById("laundry-panel-analysis"),
  laundryExportExcel: document.getElementById("laundry-export-excel"),
  laundryRows: document.getElementById("laundry-rows"),
  laundryMobileCards: document.getElementById("laundry-mobile-cards"),
  laundryCount: document.getElementById("laundry-count"),
  laundryDbStatus: document.getElementById("laundry-db-status"),
  laundryMissingWarning: document.getElementById("laundry-missing-warning"),
  laundryFilterProperty: document.getElementById("laundry-filter-property"),
  laundryFilterDateFrom: document.getElementById("laundry-filter-date-from"),
  laundryFilterDateTo: document.getElementById("laundry-filter-date-to"),
  laundryFilterSearch: document.getElementById("laundry-filter-search"),
  laundryResumeDateField: document.getElementById("laundry-resume-date-field"),
  laundryResumeDetail: document.getElementById("laundry-resume-detail"),
  laundryResumeFilterProperty: document.getElementById("laundry-resume-filter-property"),
  laundryResumeFilterDateFrom: document.getElementById("laundry-resume-filter-date-from"),
  laundryResumeFilterDateTo: document.getElementById("laundry-resume-filter-date-to"),
  laundryResumeCount: document.getElementById("laundry-resume-count"),
  laundryResumeBody: document.getElementById("laundry-resume-body"),
  laundryAnalysisDateField: document.getElementById("laundry-analysis-date-field"),
  laundryAnalysisFilterProperty: document.getElementById("laundry-analysis-filter-property"),
  laundryAnalysisFilterDateFrom: document.getElementById("laundry-analysis-filter-date-from"),
  laundryAnalysisFilterDateTo: document.getElementById("laundry-analysis-filter-date-to"),
  laundryAnalysisStatus: document.getElementById("laundry-analysis-status"),
  laundryAnalysisChart: document.getElementById("laundry-analysis-chart"),
  laundryAnalysisLegend: document.getElementById("laundry-analysis-legend"),
  laundryEditorModal: document.getElementById("laundry-editor-modal"),
  laundryCloseModal: document.getElementById("laundry-close-modal"),
  laundryStatus: document.getElementById("laundry-status"),
  laundryProperty: document.getElementById("laundry-property"),
  laundryDate: document.getElementById("laundry-date"),
  laundryReceiveDate: document.getElementById("laundry-receive-date"),
  laundryReceivedWeight: document.getElementById("laundry-received-weight"),
  laundryNotes: document.getElementById("laundry-notes"),
  laundrySentItemsGrid: document.getElementById("laundry-sent-items-grid"),
  laundryReceivedItemsGrid: document.getElementById("laundry-received-items-grid"),
  laundrySentWeight: document.getElementById("laundry-sent-weight"),
  laundryReceivedComputedWeight: document.getElementById("laundry-received-computed-weight"),
  laundryMatchDate: document.getElementById("laundry-match-date"),
  laundryDifferenceSummary: document.getElementById("laundry-difference-summary"),
  laundrySave: document.getElementById("laundry-save"),
  laundrySaveSettings: document.getElementById("laundry-save-settings"),
  laundryAddItemType: document.getElementById("laundry-add-item-type"),
  laundryPricePerKg: document.getElementById("laundry-price-per-kg"),
  laundryEmailRecipients: document.getElementById("laundry-email-recipients"),
  laundryEmailEnabled: document.getElementById("laundry-email-enabled"),
  laundryEmailTime: document.getElementById("laundry-email-time"),
  laundryTestEmail: document.getElementById("laundry-test-email"),
  laundryManagementEmailRecipients: document.getElementById("laundry-management-email-recipients"),
  laundryManagementEmailEnabled: document.getElementById("laundry-management-email-enabled"),
  laundryManagementEmailTime: document.getElementById("laundry-management-email-time"),
  laundryManagementTestEmail: document.getElementById("laundry-management-test-email"),
  laundryItemTypesBody: document.getElementById("laundry-item-types-body"),
  laundrySettingsStatus: document.getElementById("laundry-settings-status"),
  bakerySaveSettings: document.getElementById("bakery-save-settings"),
  bakerySettingsTableTab: document.getElementById("bakery-settings-table-tab"),
  bakerySettingsTypesTab: document.getElementById("bakery-settings-types-tab"),
  bakerySettingsTablePanel: document.getElementById("bakery-settings-table-panel"),
  bakerySettingsTypesPanel: document.getElementById("bakery-settings-types-panel"),
  bakerySelectedBase: document.getElementById("bakery-selected-base"),
  bakeryHostelCapacity: document.getElementById("bakery-hostel-capacity"),
  bakerySettingsEmailRecipients: document.getElementById("bakery-settings-email-recipients"),
  bakeryBreadTableBody: document.getElementById("bakery-bread-table-body"),
  bakeryBreadTypesBody: document.getElementById("bakery-bread-types-body"),
  bakeryAddBreadType: document.getElementById("bakery-add-bread-type"),
  bakeryBreadTypesTotal: document.getElementById("bakery-bread-types-total"),
  bakerySettingsStatus: document.getElementById("bakery-settings-status"),
  reviewsScreenList: document.getElementById("reviews-screen-list"),
  reviewsScreenResume: document.getElementById("reviews-screen-resume"),
  reviewsScreenRating: document.getElementById("reviews-screen-rating"),
  reviewsScreenPanelList: document.getElementById("reviews-screen-panel-list"),
  reviewsScreenPanelResume: document.getElementById("reviews-screen-panel-resume"),
  reviewsScreenPanelRating: document.getElementById("reviews-screen-panel-rating"),
  reviewsPropertyFilter: document.getElementById("reviews-property-filter"),
  reviewsSourceFilter: document.getElementById("reviews-source-filter"),
  reviewsSearch: document.getElementById("reviews-search"),
  reviewsFromDate: document.getElementById("reviews-from-date"),
  reviewsToDate: document.getElementById("reviews-to-date"),
  reviewsScoreFrom: document.getElementById("reviews-score-from"),
  reviewsScoreTo: document.getElementById("reviews-score-to"),
  reviewsCount: document.getElementById("reviews-count"),
  reviewsPagination: document.getElementById("reviews-pagination"),
  reviewsPrevPage: document.getElementById("reviews-prev-page"),
  reviewsNextPage: document.getElementById("reviews-next-page"),
  reviewsPageStatus: document.getElementById("reviews-page-status"),
  reviewsKpiAverage12m: document.getElementById("reviews-kpi-average-12m"),
  reviewsKpiAverageYear: document.getElementById("reviews-kpi-average-year"),
  reviewsKpiAverageLastMonth: document.getElementById("reviews-kpi-average-last-month"),
  reviewsKpiAverageThisMonth: document.getElementById("reviews-kpi-average-this-month"),
  reviewsRows: document.getElementById("reviews-rows"),
  reviewsMobileCards: document.getElementById("reviews-mobile-cards"),
  reviewsResumeRows: document.getElementById("reviews-resume-rows"),
  reviewsResumeStatus: document.getElementById("reviews-resume-status"),
  reviewDetailModal: document.getElementById("review-detail-modal"),
  reviewDetailClose: document.getElementById("review-detail-close"),
  reviewsDetail: document.getElementById("reviews-detail"),
  reviewsStatus: document.getElementById("reviews-status"),
  reviewsQaPrompt: document.getElementById("reviews-qa-prompt"),
  reviewsQaSubmit: document.getElementById("reviews-qa-submit"),
  reviewsQaStatus: document.getElementById("reviews-qa-status"),
  reviewsQaAnswer: document.getElementById("reviews-qa-answer"),
  reviewsAnalysisChart: document.getElementById("reviews-analysis-chart"),
  reviewsAnalysisLegend: document.getElementById("reviews-analysis-legend"),
  reviewsAnalysisStatus: document.getElementById("reviews-analysis-status"),
  reviewsExport: document.getElementById("reviews-export"),
  reviewsRefresh: document.getElementById("reviews-refresh"),
  reviewsImportStatus: document.getElementById("reviews-import-status"),
  reviewsImportProperty: document.getElementById("reviews-import-property"),
  reviewsImportSource: document.getElementById("reviews-import-source"),
  reviewsImportKind: document.getElementById("reviews-import-kind"),
  reviewsImportFiles: document.getElementById("reviews-import-files"),
  reviewsImportDropzone: document.getElementById("reviews-import-dropzone"),
  reviewsImportFileSummary: document.getElementById("reviews-import-file-summary"),
  reviewsBrowseFiles: document.getElementById("reviews-browse-files"),
  reviewsParseUpload: document.getElementById("reviews-parse-upload"),
  reviewsConfirmImport: document.getElementById("reviews-confirm-import"),
  reviewsStagingCount: document.getElementById("reviews-staging-count"),
  reviewsStagingRows: document.getElementById("reviews-staging-rows"),
  reviewsLastDatesBody: document.getElementById("reviews-last-dates-body"),
  reviewsImportRuns: document.getElementById("reviews-import-runs"),
  reviewsPropertiesBody: document.getElementById("reviews-properties-body"),
  reviewsPropertiesStatus: document.getElementById("reviews-properties-status"),
  reviewsAddProperty: document.getElementById("reviews-add-property"),
  reviewsSourcesBody: document.getElementById("reviews-sources-body"),
  reviewsSourcesStatus: document.getElementById("reviews-sources-status"),
  reviewsSaveSources: document.getElementById("reviews-save-sources"),
  reviewsGoogleStatus: document.getElementById("reviews-google-status"),
  reviewsGoogleConnect: document.getElementById("reviews-google-connect"),
  reviewsGoogleLoadLocations: document.getElementById("reviews-google-load-locations"),
  reviewsGoogleSaveMapping: document.getElementById("reviews-google-save-mapping"),
  reviewsGoogleSync: document.getElementById("reviews-google-sync"),
  reviewsGoogleMappingsBody: document.getElementById("reviews-google-mappings-body"),
};

init().catch((e) => {
  console.error(e);
  setDbStatus("Failed to initialize app.");
});

async function init() {
  ensureToastHost();
  resetSortDefault();
  bindEvents();
  await initAuth();
  await loadAccess();
  if (!canApp("communications") && canApp("guests")) state.currentView = "guests";
  else if (!canApp("communications") && !canApp("guests") && canApp("cash")) state.currentView = "cash";
  else if (!canApp("communications") && !canApp("guests") && !canApp("cash") && canApp("lost-found")) state.currentView = "lost-found";
  else if (!canApp("communications") && !canApp("guests") && !canApp("cash") && !canApp("lost-found") && canApp("groups")) state.currentView = "groups";
  else if (!canApp("communications") && !canApp("guests") && !canApp("cash") && !canApp("lost-found") && !canApp("groups") && canApp("services")) state.currentView = "services";
  else if (!canApp("communications") && !canApp("guests") && !canApp("cash") && !canApp("lost-found") && !canApp("groups") && !canApp("services") && canApp("shopping")) state.currentView = "shopping";
  else if (!canApp("communications") && !canApp("guests") && !canApp("cash") && !canApp("lost-found") && !canApp("groups") && !canApp("services") && !canApp("shopping") && canApp("hours")) state.currentView = "hours";
  else if (!canApp("communications") && !canApp("guests") && !canApp("cash") && !canApp("lost-found") && !canApp("groups") && !canApp("services") && !canApp("shopping") && !canApp("hours") && canApp("bakery")) state.currentView = "bakery";
  else if (!canApp("communications") && !canApp("guests") && !canApp("cash") && !canApp("lost-found") && !canApp("groups") && !canApp("services") && !canApp("shopping") && !canApp("hours") && !canApp("bakery") && canApp("laundry")) state.currentView = "laundry";
  else if (!canApp("communications") && !canApp("guests") && !canApp("cash") && !canApp("lost-found") && !canApp("groups") && !canApp("services") && !canApp("shopping") && !canApp("hours") && !canApp("bakery") && !canApp("laundry") && canApp("reviews")) state.currentView = "reviews";
  else if (!canApp("communications") && state.access.settingsFeatures.length > 0) state.currentView = "settings";
  applyInitialRouteFromUrl();
  if (!canAccessGeneralSettings() && canSettings("guests")) state.settingsSection = "guests";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && canSettings("cash")) state.settingsSection = "cash";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && !canSettings("cash") && canSettings("reviews")) state.settingsSection = "reviews";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && !canSettings("cash") && !canSettings("reviews") && canSettings("groups")) state.settingsSection = "groups";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && !canSettings("cash") && !canSettings("reviews") && !canSettings("groups") && canSettings("services")) state.settingsSection = "services";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && !canSettings("cash") && !canSettings("reviews") && !canSettings("groups") && !canSettings("services") && canSettings("shopping")) state.settingsSection = "shopping";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && !canSettings("cash") && !canSettings("reviews") && !canSettings("groups") && !canSettings("services") && !canSettings("shopping") && canSettings("hours")) state.settingsSection = "hours";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && !canSettings("cash") && !canSettings("reviews") && !canSettings("groups") && !canSettings("services") && !canSettings("shopping") && !canSettings("hours") && canSettings("bakery")) state.settingsSection = "bakery";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && !canSettings("cash") && !canSettings("reviews") && !canSettings("groups") && !canSettings("services") && !canSettings("shopping") && !canSettings("hours") && !canSettings("bakery") && canSettings("laundry")) state.settingsSection = "laundry";
  else if (!canAccessGeneralSettings() && !canSettings("guests") && !canSettings("cash") && canSettings("admin-users")) state.settingsSection = "admin-users";
  renderLayout();
  renderSettingsSection();
  renderCategoryFilterOptions();
  renderReviewPropertyOptions();
  render();
  if (canApp("communications")) loadSidebarReviewSummary({ silent: true }).catch(() => {});
  await ensureCurrentViewData();
  if (canApp("guests")) loadGuestsData({ silent: true }).then(() => renderLayout()).catch(() => {});
  if (canApp("cash")) loadCashData({ silent: true }).then(() => renderLayout()).catch(() => {});
  if (canApp("shopping")) loadShoppingData({ silent: true }).then(() => renderLayout()).catch(() => {});
  if (canApp("hours")) loadHoursData({ silent: true }).catch(() => {});
  if (canApp("bakery")) loadBakeryData({ silent: true }).then(() => renderLayout()).catch(() => {});
  if (canApp("laundry")) loadLaundryRecords({ silent: true }).then(() => renderLayout()).catch(() => {});
  startAutoRefresh();
}

function bindEvents() {
  els.navCommunications.addEventListener("click", () => setView("communications"));
  els.navGuests?.addEventListener("click", () => setView("guests"));
  els.navCash?.addEventListener("click", () => setView("cash"));
  els.navLostFound.addEventListener("click", () => setView("lost-found"));
  els.navReviews.addEventListener("click", () => setView("reviews"));
  els.navGroups.addEventListener("click", () => setView("groups"));
  els.navServices.addEventListener("click", () => setView("services"));
  els.navShopping.addEventListener("click", () => setView("shopping"));
  els.navHours.addEventListener("click", () => setView("hours"));
  els.navBakery.addEventListener("click", () => setView("bakery"));
  els.navLaundry.addEventListener("click", () => setView("laundry"));
  els.sidebarReviewSummaryCard?.addEventListener("click", async () => {
    state.reviewScreen = "resume";
    await setView("reviews");
    setReviewScreen("resume");
  });
  els.mobileMenuToggle?.addEventListener("click", toggleMobileNav);
  els.openSettings.addEventListener("click", () => setView("settings"));
  els.closeSettingsGeneral?.addEventListener("click", () => setView("communications"));
  els.closeSettings.addEventListener("click", () => setView("communications"));
  els.closeSettingsCash?.addEventListener("click", () => setView("cash"));
  els.closeSettingsAdmin.addEventListener("click", () => setView("communications"));
  els.closeSettingsReviews.addEventListener("click", () => setView("reviews"));
  els.closeSettingsGroups.addEventListener("click", () => setView("groups"));
  els.closeSettingsServices.addEventListener("click", () => setView("services"));
  els.closeSettingsShopping.addEventListener("click", () => setView("shopping"));
  els.closeSettingsHours?.addEventListener("click", () => setView("hours"));
  els.closeSettingsBakery.addEventListener("click", () => setView("bakery"));
  els.closeSettingsLaundry?.addEventListener("click", () => setView("laundry"));
  els.closeSettingsGuests?.addEventListener("click", () => setView("guests"));
  els.settingsMenuGeneral?.addEventListener("click", () => setSettingsSection("general"));
  els.settingsMenuCommunications.addEventListener("click", () => setSettingsSection("communications"));
  els.settingsMenuGuests?.addEventListener("click", () => setSettingsSection("guests"));
  els.settingsMenuCash?.addEventListener("click", () => setSettingsSection("cash"));
  els.settingsMenuReviews.addEventListener("click", () => setSettingsSection("reviews"));
  els.settingsMenuGroups.addEventListener("click", () => setSettingsSection("groups"));
  els.settingsMenuServices.addEventListener("click", () => setSettingsSection("services"));
  els.settingsMenuShopping.addEventListener("click", () => setSettingsSection("shopping"));
  els.settingsMenuHours?.addEventListener("click", () => setSettingsSection("hours"));
  els.settingsMenuBakery.addEventListener("click", () => setSettingsSection("bakery"));
  els.settingsMenuLaundry.addEventListener("click", () => setSettingsSection("laundry"));
  els.settingsMenuAdminUsers.addEventListener("click", () => setSettingsSection("admin-users"));
  els.shoppingTabCurrent.addEventListener("click", () => setShoppingTab("current"));
  els.shoppingTabHistory.addEventListener("click", () => setShoppingTab("history"));
  els.shoppingNewOrder.addEventListener("click", createShoppingOrder);
  els.shoppingSaveOrder.addEventListener("click", () => saveShoppingOrderDraft(false));
  els.shoppingSubmitOrder.addEventListener("click", submitShoppingOrder);
  els.shoppingFilterCategory?.addEventListener("change", onShoppingFilterChange);
  els.shoppingFilterStored?.addEventListener("change", onShoppingFilterChange);
  els.shoppingGroupBy?.addEventListener("change", onShoppingFilterChange);
  els.shoppingOpenRows.addEventListener("input", onShoppingOrderInput);
  els.shoppingOpenRows.addEventListener("change", onShoppingOrderInput);
  els.shoppingMobileCards?.addEventListener("input", onShoppingOrderInput);
  els.shoppingMobileCards?.addEventListener("change", onShoppingOrderInput);
  els.shoppingHistoryDateFrom?.addEventListener("input", onShoppingHistoryFilterChange);
  els.shoppingHistoryDateTo?.addEventListener("input", onShoppingHistoryFilterChange);
  els.shoppingHistoryName?.addEventListener("input", onShoppingHistoryFilterChange);
  els.shoppingHistoryCategory?.addEventListener("change", onShoppingHistoryFilterChange);
  els.shoppingHistorySupplier?.addEventListener("change", onShoppingHistoryFilterChange);
  els.shoppingHistoryRows.addEventListener("click", onShoppingHistoryAction);
  els.shoppingHistoryMobileCards?.addEventListener("click", onShoppingHistoryAction);
  els.shoppingExportExcel?.addEventListener("click", exportShoppingDetailToExcel);
  els.shoppingExportPdf?.addEventListener("click", exportShoppingDetailToPdf);
  els.shoppingCopyOrder?.addEventListener("click", copyShoppingOrderAsDraft);
  els.shoppingDetailClose.addEventListener("click", closeShoppingDetailModal);
  els.shoppingReopenOrder.addEventListener("click", reopenLatestShoppingOrder);
  els.shoppingSaveSettings.addEventListener("click", saveShoppingSettings);
  els.shoppingAddItem.addEventListener("click", addShoppingSettingItem);
  els.shoppingSettingsItemsBody.addEventListener("input", onShoppingSettingsInput);
  els.shoppingSettingsItemsBody.addEventListener("change", onShoppingSettingsInput);
  els.shoppingSettingsItemsBody.addEventListener("click", onShoppingSettingsAction);
  els.shoppingSettingsCategoryColors?.addEventListener("input", onShoppingSettingsInput);
  els.shoppingSettingsCategoryColors?.addEventListener("change", onShoppingSettingsInput);
  els.shoppingSettingsWeekdays?.addEventListener("change", onShoppingSettingsAction);
  els.guestsTabList?.addEventListener("click", () => setGuestsScreen("list"));
  els.guestsTabDescriptions?.addEventListener("click", () => setGuestsScreen("descriptions"));
  els.guestsTabBlacklist?.addEventListener("click", () => setGuestsScreen("blacklist"));
  els.guestsSettingsConfigTab?.addEventListener("click", () => setGuestsSettingsTab("config"));
  els.guestsSettingsSefTab?.addEventListener("click", () => setGuestsSettingsTab("sef"));
  els.guestsSettingsApiTab?.addEventListener("click", () => setGuestsSettingsTab("api"));
  els.guestsShowActive?.addEventListener("change", onGuestsFilterInput);
  els.guestsFilterHa?.addEventListener("change", onGuestsFilterInput);
  [
    els.guestsFilterSearch,
    els.guestsFilterNationality,
    els.guestsFilterCheckinFrom,
    els.guestsFilterCheckinTo,
    els.guestsFilterCheckoutFrom,
    els.guestsFilterCheckoutTo,
  ].forEach((el) => el?.addEventListener("input", onGuestsFilterInput));
  els.guestsRows?.addEventListener("click", onGuestsAction);
  els.guestsRows?.addEventListener("input", onGuestsDraftInput);
  els.guestsRows?.addEventListener("change", onGuestsDraftInput);
  els.guestsRows?.addEventListener("focusin", releaseGuestsIssuerCountryGuard);
  els.guestsRows?.addEventListener("click", onGuestsQuickEditClick);
  els.guestsRows?.addEventListener("change", onGuestsQuickEditChange);
  els.guestsRows?.addEventListener("keydown", onGuestsKeydown);
  els.guestsMobileCards?.addEventListener("click", onGuestsAction);
  els.guestsMobileCards?.addEventListener("input", onGuestsDraftInput);
  els.guestsMobileCards?.addEventListener("change", onGuestsDraftInput);
  els.guestsMobileCards?.addEventListener("focusin", releaseGuestsIssuerCountryGuard);
  els.guestsMobileCards?.addEventListener("click", onGuestsQuickEditClick);
  els.guestsMobileCards?.addEventListener("change", onGuestsQuickEditChange);
  els.guestsMobileCards?.addEventListener("keydown", onGuestsKeydown);
  els.guestsDescriptionsRows?.addEventListener("focusin", onGuestDescriptionFocusIn);
  els.guestsDescriptionsRows?.addEventListener("input", onGuestDescriptionInput);
  els.guestsDescriptionsRows?.addEventListener("focusout", onGuestDescriptionFocusOut);
  els.guestsDescriptionsMobileCards?.addEventListener("focusin", onGuestDescriptionFocusIn);
  els.guestsDescriptionsMobileCards?.addEventListener("input", onGuestDescriptionInput);
  els.guestsDescriptionsMobileCards?.addEventListener("focusout", onGuestDescriptionFocusOut);
  els.guestsDescriptionsFilterRoom?.addEventListener("input", onGuestDescriptionsFilterInput);
  els.guestsDescriptionsFilterDescription?.addEventListener("input", onGuestDescriptionsFilterInput);
  els.guestsBlacklistRows?.addEventListener("click", onGuestsBlacklistAction);
  els.guestsBlacklistRows?.addEventListener("input", onGuestsBlacklistDraftInput);
  els.guestsBlacklistRows?.addEventListener("change", onGuestsBlacklistDraftInput);
  els.guestsBlacklistMobileCards?.addEventListener("click", onGuestsBlacklistAction);
  els.guestsBlacklistMobileCards?.addEventListener("input", onGuestsBlacklistDraftInput);
  els.guestsBlacklistMobileCards?.addEventListener("change", onGuestsBlacklistDraftInput);
  [els.guestsBlacklistFilterSearch, els.guestsBlacklistFilterReported, els.guestsBlacklistFilterNationality].forEach((el) => el?.addEventListener("input", onGuestsBlacklistFilterInput));
  els.guestsExportExcel?.addEventListener("click", exportGuestsToExcel);
  els.guestsSendPending?.addEventListener("click", sendPendingGuests);
  els.guestsSaveSettings?.addEventListener("click", saveGuestsSettings);
  els.guestsSettingsSendTime?.addEventListener("input", onGuestsSettingsInput);
  els.guestsSettingsMappingBody?.addEventListener("change", onGuestsSettingsInput);
  els.cashRows?.addEventListener("click", onCashTableAction);
  els.cashRows?.addEventListener("input", onCashTableInput);
  els.cashMobileCards?.addEventListener("click", onCashTableAction);
  els.cashMobileCards?.addEventListener("input", onCashTableInput);
  els.cashFilterDateFrom?.addEventListener("input", onCashFilterInput);
  els.cashFilterDateTo?.addEventListener("input", onCashFilterInput);
  els.cashFilterShift?.addEventListener("input", onCashFilterInput);
  els.cashFilterName?.addEventListener("input", onCashFilterInput);
  els.cashTabList?.addEventListener("click", () => {
    state.cashScreen = "list";
    renderCash();
  });
  els.cashTabDetail?.addEventListener("click", () => {
    state.cashScreen = "detail";
    renderCash();
  });
  els.cashTabItems?.addEventListener("click", () => {
    state.cashScreen = "items";
    renderCash();
  });
  els.cashTabResume?.addEventListener("click", () => {
    state.cashScreen = "resume";
    renderCash();
  });
  els.cashSettingsConfigTab?.addEventListener("click", () => setCashSettingsTab("config"));
  els.cashSettingsMinTab?.addEventListener("click", () => setCashSettingsTab("min"));
  els.cashSaveSettings?.addEventListener("click", saveCashSettings);
  els.cashAddShift?.addEventListener("click", addCashShiftSetting);
  els.cashAddItem?.addEventListener("click", addCashItemSetting);
  els.cashMoneyClose?.addEventListener("click", closeCashMoneyModal);
  els.cashMoneyBody?.addEventListener("input", onCashMoneyModalInput);
  els.cashMoneySave?.addEventListener("click", saveCashMoneyModal);
  els.cashSettingsShiftsBody?.addEventListener("input", onCashSettingsInput);
  els.cashSettingsItemsBody?.addEventListener("input", onCashSettingsInput);
  els.cashSettingsMinBody?.addEventListener("input", onCashSettingsInput);
  els.cashSettingsManagerAlertEmail?.addEventListener("input", onCashSettingsInput);
  els.cashSettingsMinimumEmailEnabled?.addEventListener("change", onCashSettingsInput);
  els.cashSettingsMaximumEmailEnabled?.addEventListener("change", onCashSettingsInput);
  els.cashSettingsMaximumCash?.addEventListener("input", onCashSettingsInput);
  els.cashSettingsShiftsBody?.addEventListener("click", onCashSettingsAction);
  els.cashSettingsItemsBody?.addEventListener("click", onCashSettingsAction);
  els.cashItemsClose?.addEventListener("click", closeCashItemsModal);
  els.cashItemsBody?.addEventListener("input", onCashItemsModalInput);
  els.cashItemsBody?.addEventListener("click", onCashItemsModalAction);
  els.cashItemsSave?.addEventListener("click", saveCashItemsModal);
  els.hoursExportExcel?.addEventListener("click", exportHoursToExcel);
  els.hoursTabList?.addEventListener("click", () => setHoursScreen("list"));
  els.hoursTabResume?.addEventListener("click", () => setHoursScreen("resume"));
  [els.hoursFilterPerson, els.hoursResumeFilterPerson].forEach((el) => el?.addEventListener("change", onHoursFilterInput));
  [els.hoursFilterDateFrom, els.hoursFilterDateTo, els.hoursResumeFilterDateFrom, els.hoursResumeFilterDateTo].forEach((el) => el?.addEventListener("input", onHoursFilterInput));
  els.hoursRows?.addEventListener("click", onHoursAction);
  els.hoursRows?.addEventListener("input", onHoursDraftInput);
  els.hoursRows?.addEventListener("change", onHoursDraftInput);
  els.hoursMobileCards?.addEventListener("click", onHoursAction);
  els.hoursMobileCards?.addEventListener("input", onHoursDraftInput);
  els.hoursMobileCards?.addEventListener("change", onHoursDraftInput);
  els.hoursSaveSettings?.addEventListener("click", saveHoursSettings);
  els.hoursSettingsPersons?.addEventListener("input", onHoursSettingsInput);
  els.shoppingSubmitName.addEventListener("input", () => {
    state.shoppingSubmitName = clean(els.shoppingSubmitName.value);
  });
  els.shoppingSubmitNotes?.addEventListener("input", () => {
    state.shoppingSubmitNotes = clean(els.shoppingSubmitNotes.value);
  });
  els.bakeryTabCurrent?.addEventListener("click", () => setBakeryTab("current"));
  els.bakeryTabHistory?.addEventListener("click", () => setBakeryTab("history"));
  els.bakeryTabResume?.addEventListener("click", () => setBakeryTab("resume"));
  els.bakeryNewOrder?.addEventListener("click", createBakeryOrder);
  els.bakerySaveOrder?.addEventListener("click", saveBakeryOrderDraft);
  els.bakerySubmitOrder?.addEventListener("click", submitBakeryOrder);
  els.bakeryOpenRows?.addEventListener("change", onBakeryOrderInput);
  els.bakeryMobileCards?.addEventListener("change", onBakeryOrderInput);
  els.bakeryHistoryRows?.addEventListener("click", onBakeryHistoryAction);
  els.bakeryHistoryMobileCards?.addEventListener("click", onBakeryHistoryAction);
  els.bakeryDetailClose?.addEventListener("click", closeBakeryDetailModal);
  els.bakeryDetailResend?.addEventListener("click", resendBakeryOrderEmail);
  els.bakerySubmitName?.addEventListener("input", () => {
    state.bakerySubmitName = clean(els.bakerySubmitName.value);
    if (state.bakeryOpenOrder && els.bakeryGeneratedText) {
      refreshBakeryOpenOrderDerivedState();
      els.bakeryGeneratedText.innerHTML = buildBakeryGeneratedHtmlClient(state.bakeryOpenOrder, state.bakerySubmitName);
    }
  });
  els.bakerySettingsTableTab?.addEventListener("click", () => setBakerySettingsTab("table"));
  els.bakerySettingsTypesTab?.addEventListener("click", () => setBakerySettingsTab("types"));
  els.bakerySaveSettings?.addEventListener("click", saveBakerySettings);
  els.generalSaveSettings?.addEventListener("click", saveSettings);
  els.bakerySelectedBase?.addEventListener("change", onBakerySettingsInput);
  els.bakeryHostelCapacity?.addEventListener("input", onBakerySettingsInput);
  els.bakerySettingsEmailRecipients?.addEventListener("input", onBakerySettingsInput);
  [els.generalEmailProvider, els.generalEmailSmtpHost, els.generalEmailSmtpPort, els.generalEmailSmtpSecure, els.generalEmailSmtpUser, els.generalEmailSmtpPassword, els.generalEmailFromEmail, els.generalEmailFromName]
    .filter(Boolean)
    .forEach((el) => el.addEventListener(el.type === "checkbox" || el.tagName === "SELECT" ? "change" : "input", onGeneralSettingsInput));
  els.bakeryBreadTableBody?.addEventListener("input", onBakerySettingsInput);
  els.bakeryBreadTypesBody?.addEventListener("input", onBakerySettingsInput);
  els.bakeryBreadTypesBody?.addEventListener("click", onBakerySettingsAction);
  els.bakeryAddBreadType?.addEventListener("click", addBakeryBreadType);
  els.laundryNew?.addEventListener("click", async () => {
    await ensureLaundryData();
    resetLaundryDraft();
    openLaundryModal();
  });
  els.laundryTabList?.addEventListener("click", () => setLaundryScreen("list"));
  els.laundryTabResume?.addEventListener("click", () => setLaundryScreen("resume"));
  els.laundryTabAnalysis?.addEventListener("click", () => setLaundryScreen("analysis"));
  els.laundryExportExcel?.addEventListener("click", exportLaundryToExcel);
  els.laundryCloseModal?.addEventListener("click", closeLaundryModal);
  els.laundryFilterProperty?.addEventListener("change", onLaundryFilterInput);
  [els.laundryFilterDateFrom, els.laundryFilterDateTo, els.laundryFilterSearch].forEach((el) => el?.addEventListener("input", onLaundryFilterInput));
  [els.laundryResumeDateField, els.laundryResumeFilterProperty, els.laundryResumeDetail].forEach((el) => el?.addEventListener("change", onLaundryResumeFilterInput));
  [els.laundryResumeFilterDateFrom, els.laundryResumeFilterDateTo].forEach((el) => el?.addEventListener("input", onLaundryResumeFilterInput));
  [els.laundryAnalysisDateField, els.laundryAnalysisFilterProperty].forEach((el) => el?.addEventListener("change", onLaundryAnalysisFilterInput));
  [els.laundryAnalysisFilterDateFrom, els.laundryAnalysisFilterDateTo].forEach((el) => el?.addEventListener("input", onLaundryAnalysisFilterInput));
  [els.laundryProperty, els.laundryDate, els.laundryReceivedWeight, els.laundryNotes].forEach((el) =>
    el?.addEventListener(el.tagName === "SELECT" ? "change" : "input", onLaundryDraftInput)
  );
  els.laundrySentItemsGrid?.addEventListener("input", onLaundryDraftGridInput);
  els.laundryReceivedItemsGrid?.addEventListener("input", onLaundryDraftGridInput);
  els.laundrySave?.addEventListener("click", saveLaundryRecord);
  els.laundryRows?.addEventListener("click", onLaundryRowClick);
  els.laundryMobileCards?.addEventListener("click", onLaundryRowClick);
  els.laundrySaveSettings?.addEventListener("click", saveLaundrySettings);
  els.laundryAddItemType?.addEventListener("click", addLaundrySettingItemType);
  els.laundryPricePerKg?.addEventListener("input", onLaundrySettingsInput);
  els.laundryEmailRecipients?.addEventListener("input", onLaundrySettingsInput);
  els.laundryEmailEnabled?.addEventListener("change", onLaundrySettingsInput);
  els.laundryEmailTime?.addEventListener("input", onLaundrySettingsInput);
  els.laundryTestEmail?.addEventListener("click", triggerLaundryEmailNow);
  els.laundryManagementEmailRecipients?.addEventListener("input", onLaundrySettingsInput);
  els.laundryManagementEmailEnabled?.addEventListener("change", onLaundrySettingsInput);
  els.laundryManagementEmailTime?.addEventListener("input", onLaundrySettingsInput);
  els.laundryManagementTestEmail?.addEventListener("click", triggerLaundryManagementEmailNow);
  els.laundryItemTypesBody?.addEventListener("input", onLaundrySettingsInput);
  els.laundryItemTypesBody?.addEventListener("click", onLaundrySettingsAction);
  els.settingsReviewsImportTab.addEventListener("click", () => setReviewSettingsScreen("import"));
  els.settingsReviewsConfigTab.addEventListener("click", () => setReviewSettingsScreen("config"));
  els.rows.addEventListener("click", onRowAction);
  els.rows.addEventListener("input", onRowDraftInput);
  els.rows.addEventListener("keydown", onRowKeydown);
  els.rows.addEventListener("change", onRowStatusToggle);
  els.communicationsMobileCards?.addEventListener("click", onRowAction);
  els.communicationsMobileCards?.addEventListener("input", onRowDraftInput);
  els.communicationsMobileCards?.addEventListener("keydown", onRowKeydown);
  els.communicationsMobileCards?.addEventListener("change", onRowStatusToggle);
  els.lostFoundRows.addEventListener("click", onLostFoundAction);
  els.lostFoundRows.addEventListener("input", onLostFoundDraftInput);
  els.lostFoundRows.addEventListener("keydown", onLostFoundKeydown);
  els.lostFoundRows.addEventListener("change", onLostFoundStatusToggle);
  els.lostFoundMobileCards?.addEventListener("click", onLostFoundAction);
  els.lostFoundMobileCards?.addEventListener("input", onLostFoundDraftInput);
  els.lostFoundMobileCards?.addEventListener("keydown", onLostFoundKeydown);
  els.lostFoundMobileCards?.addEventListener("change", onLostFoundStatusToggle);
  els.tableHead.addEventListener("click", onSortToggle);
  els.resetSort.addEventListener("click", () => {
    resetSortDefault();
    render();
    showToast("Default sort applied: Date/Time newest first.", "info");
  });
  [els.search, els.showActive, els.groupCommunications, els.statusFilter, els.categoryFilter, els.fromDate, els.toDate].forEach((el) =>
    el.addEventListener("input", render)
  );
  ["focus", "pointerdown", "touchstart"].forEach((eventName) => {
    els.search?.addEventListener(eventName, releaseCommunicationsSearchGuard, { passive: true });
  });
  els.showActive.addEventListener("change", render);
  els.groupCommunications.addEventListener("change", render);
  [
    els.lostFoundOnlyOpen,
    els.lostFoundFilterNumber,
    els.lostFoundFilterDate,
    els.lostFoundFilterWhoFound,
    els.lostFoundFilterWhoRecorded,
    els.lostFoundFilterWhere,
    els.lostFoundFilterObject,
    els.lostFoundFilterNotes,
    els.lostFoundFilterStored,
  ].forEach((el) => el.addEventListener("input", renderLostFound));
  els.lostFoundOnlyOpen.addEventListener("change", renderLostFound);
  els.lostFoundFilterStored.addEventListener("change", renderLostFound);
  els.excelInput.addEventListener("change", importFromExcel);
  els.exportCsv.addEventListener("click", exportToCsv);
  els.authLogout.addEventListener("click", signOut);
  els.addCategory.addEventListener("click", addCategory);
  els.settingsCategoriesBody.addEventListener("click", removeCategoryClick);
  els.settingsCategoriesBody.addEventListener("input", settingsCategoryInput);
  [els.settingEmailEnabled, els.settingEmailFrequency, els.settingEmailTime, els.settingEmailRecipients, els.settingEmailFrequency2, els.settingEmailTime2, els.settingEmailRecipients2].forEach((el) =>
    el.addEventListener("input", updateEmailSettings)
  );
  els.testEmailNow.addEventListener("click", triggerEmailNow);
  els.saveSettings.addEventListener("click", saveSettings);
  els.groupsNew.addEventListener("click", async () => {
    await refreshGroupSettingsForEditor();
    resetGroupDraft();
    openGroupModal();
  });
  els.groupsTabList?.addEventListener("click", () => setGroupsScreen("list"));
  els.groupsTabResume?.addEventListener("click", () => setGroupsScreen("resume"));
  els.groupCloseModal.addEventListener("click", closeGroupModal);
  els.groupTabDetails.addEventListener("click", () => setGroupEditorTab("details"));
  els.groupTabEmail.addEventListener("click", () => setGroupEditorTab("email"));
  els.groupTabConfirmation.addEventListener("click", () => setGroupEditorTab("confirmation"));
  els.groupTabFinalConfirmation.addEventListener("click", () => setGroupEditorTab("final-confirmation"));
  els.groupProposalLanguage.addEventListener("change", () => {
    state.groupProposalLanguage = normalizeProposalLanguage(els.groupProposalLanguage.value);
    state.groupDraft.language = state.groupProposalLanguage;
    renderGroupProposalEmail();
  });
  els.groupCopyEmail.addEventListener("click", copyGroupEmailText);
  [els.groupReservationNumber, els.groupName, els.groupEmail, els.groupCheckIn, els.groupCheckOut, els.groupGuests, els.groupOptionDate, els.groupStatusField, els.groupObservation].forEach((el) =>
    el.addEventListener("input", onGroupDraftInput)
  );
  els.groupCheckOut.addEventListener("focus", prepareGroupCheckOutPicker);
  els.groupCheckOut.addEventListener("pointerdown", prepareGroupCheckOutPicker);
  els.groupAddRoomItem.addEventListener("click", addGroupRoomItem);
  els.groupRoomItemsBody.addEventListener("input", onGroupRoomItemInput);
  els.groupRoomItemsBody.addEventListener("change", onGroupRoomItemInput);
  els.groupRoomItemsBody.addEventListener("click", onGroupRoomItemAction);
  els.groupSave.addEventListener("click", saveGroupProposal);
  els.groupDelete.addEventListener("click", deleteGroupProposal);
  els.groupsExportExcel.addEventListener("click", exportGroupsToExcel);
  els.groupsExportPdf.addEventListener("click", exportGroupsToPdf);
  els.groupsShowActive.addEventListener("change", () => {
    state.groupsShowActive = els.groupsShowActive.checked;
    renderGroups();
  });
  els.groupsResumeMonthMode?.addEventListener("change", () => {
    state.groupResumeMonthMode = clean(els.groupsResumeMonthMode.value) === "checkin" ? "checkin" : "created";
    renderGroups();
  });
  [els.groupsFilterCreatedFrom, els.groupsFilterCreatedTo, els.groupsFilterDateFrom, els.groupsFilterDateTo, els.groupsFilterSearch].forEach((el) =>
    el?.addEventListener("input", onGroupFilterInput)
  );
  els.groupsRows.addEventListener("click", onGroupRowClick);
  els.groupsMobileCards?.addEventListener("click", onGroupRowClick);
  els.groupsRows.closest("table").querySelector("thead").addEventListener("click", onGroupSortToggle);
  els.groupsSettingsConfigTab.addEventListener("click", () => setGroupSettingsTab("config"));
  els.groupsSettingsProposalTab.addEventListener("click", () => setGroupSettingsTab("proposal"));
  els.groupsSettingsConfirmationTab.addEventListener("click", () => setGroupSettingsTab("confirmation"));
  els.groupsSettingsFinalConfirmationTab.addEventListener("click", () => setGroupSettingsTab("final-confirmation"));
  els.groupsAddRoomType.addEventListener("click", addGroupSettingsRoomType);
  els.groupsRoomTypesBody.addEventListener("input", onGroupSettingsInput);
  els.groupsRoomTypesBody.addEventListener("click", onGroupSettingsRoomTypeAction);
  els.groupsDepositPercentage.addEventListener("input", onGroupSettingsInput);
  els.groupsLastPaymentDays.addEventListener("input", onGroupSettingsInput);
  els.groupsEmailTemplate.addEventListener("input", onGroupSettingsInput);
  els.groupsConfirmationTemplate.addEventListener("input", onGroupSettingsInput);
  els.groupsFinalConfirmationTemplate.addEventListener("input", onGroupSettingsInput);
  els.groupsSaveSettings.addEventListener("click", saveGroupSettings);
  els.groupsSaveSettingsProposal.addEventListener("click", saveGroupSettings);
  els.groupsSaveSettingsConfirmation.addEventListener("click", saveGroupSettings);
  els.groupsSaveSettingsFinalConfirmation.addEventListener("click", saveGroupSettings);
  els.servicesNew.addEventListener("click", async () => {
    await ensureServicesData();
    resetServiceDraft();
    openServiceModal();
  });
  els.servicesTabList?.addEventListener("click", () => setServicesScreen("list"));
  els.servicesTabResume?.addEventListener("click", () => setServicesScreen("resume"));
  els.serviceTabDetails.addEventListener("click", () => setServiceEditorTab("details"));
  els.serviceTabConfirmation.addEventListener("click", () => setServiceEditorTab("confirmation"));
  els.serviceConfirmationLanguage.addEventListener("input", () => {
    state.serviceConfirmationLanguage = normalizeProposalLanguage(els.serviceConfirmationLanguage.value);
    state.serviceDraft.language = state.serviceConfirmationLanguage;
    renderServiceConfirmationPreview();
  });
  els.servicesRows.addEventListener("click", onServiceRowClick);
  els.servicesMobileCards?.addEventListener("click", onServiceRowClick);
  els.servicesRows.addEventListener("change", onInlineServiceStatusChange);
  els.servicesMobileCards?.addEventListener("change", onInlineServiceStatusChange);
  els.serviceCloseModal.addEventListener("click", closeServiceModal);
  [els.serviceType, els.serviceStatus, els.serviceCustomerName, els.serviceCustomerEmail, els.serviceCustomerPhone, els.servicePax, els.serviceDate, els.serviceTime, els.servicePickupLocation, els.serviceDropoffLocation, els.serviceFlightNumber, els.serviceHasReturn, els.servicePrice, els.serviceNotes, els.serviceReturnPickup, els.serviceReturnDropoff, els.serviceReturnDate, els.serviceReturnTime, els.serviceReturnFlight].forEach((el) =>
    el.addEventListener("input", onServiceDraftInput)
  );
  [els.serviceDatePicker, els.serviceReturnDatePicker].forEach((el) =>
    el.addEventListener("input", onServiceDatePickerInput)
  );
  els.serviceSave.addEventListener("click", saveService);
  els.serviceDelete.addEventListener("click", deleteService);
  els.serviceCopyConfirmation.addEventListener("click", copyServiceConfirmationText);
  els.servicesExportExcel.addEventListener("click", exportServicesToExcel);
  [els.servicesShowActive, els.servicesFilterCreatedFrom, els.servicesFilterCreatedTo, els.servicesFilterDateFrom, els.servicesFilterDateTo, els.servicesFilterName].forEach((el) =>
    el.addEventListener("input", onServiceFilterInput)
  );
  els.servicesSettingsConfigTab.addEventListener("click", () => setServiceSettingsTab("config"));
  els.servicesSettingsConfirmationTab.addEventListener("click", () => setServiceSettingsTab("confirmation"));
  els.servicesConfigsBody.addEventListener("input", onServiceSettingsInput);
  els.servicesAutomaticEmailRecipients.addEventListener("input", onServiceSettingsInput);
  els.servicesLiveFlightStatusEnabled?.addEventListener("input", onServiceSettingsInput);
  [els.servicesPriceOneWay13, els.servicesPriceOneWay47, els.servicesPriceOneWay811, els.servicesPriceOneWay1216, els.servicesPriceReturn13, els.servicesPriceReturn47, els.servicesPriceReturn811, els.servicesPriceReturn1216].forEach((el) =>
    el.addEventListener("input", onServiceSettingsInput)
  );
  els.servicesTemplateServiceType.addEventListener("input", onServiceSettingsTemplateChange);
  els.servicesTemplateLanguage.addEventListener("input", () => {
    state.serviceSettingsTemplateLanguage = normalizeProposalLanguage(els.servicesTemplateLanguage.value);
    renderServiceSettingsTemplatePreview();
  });
  els.servicesConfirmationTemplate.addEventListener("input", onServiceSettingsInput);
  els.servicesSaveSettings.addEventListener("click", saveServiceSettings);
  els.servicesSaveSettingsConfirmation.addEventListener("click", saveServiceSettings);
  els.adminCreateUser.addEventListener("click", createAdminUser);
  els.adminRefreshUsers.addEventListener("click", () => loadAdminUsers(true));
  els.reviewsScreenList.addEventListener("click", () => setReviewScreen("list"));
  els.reviewsScreenResume.addEventListener("click", () => setReviewScreen("resume"));
  els.reviewsScreenRating.addEventListener("click", () => setReviewScreen("rating"));
  els.reviewsQaSubmit.addEventListener("click", submitReviewQuestion);
  [els.reviewsPropertyFilter, els.reviewsSourceFilter, els.reviewsSearch, els.reviewsFromDate, els.reviewsToDate, els.reviewsScoreFrom, els.reviewsScoreTo].forEach((el) =>
    el.addEventListener("input", onReviewFilterInput)
  );
  els.reviewsRefresh.addEventListener("click", async () => {
    await loadReviews({ useFilters: true });
    render();
  });
  els.reviewsRows.addEventListener("click", onReviewRowClick);
  els.reviewsMobileCards?.addEventListener("click", onReviewRowClick);
  els.reviewDetailClose.addEventListener("click", closeReviewDetailModal);
  els.reviewsExport.addEventListener("click", exportReviewsToCsv);
  els.reviewsPrevPage.addEventListener("click", () => setReviewListPage(state.reviewListPage - 1));
  els.reviewsNextPage.addEventListener("click", () => setReviewListPage(state.reviewListPage + 1));
  els.reviewsParseUpload.addEventListener("click", parseReviewUploads);
  els.reviewsImportFiles.addEventListener("change", renderReviewImportFileSummary);
  els.reviewsBrowseFiles.addEventListener("click", () => els.reviewsImportFiles.click());
  els.reviewsImportDropzone.addEventListener("click", () => els.reviewsImportDropzone.focus());
  els.reviewsImportDropzone.addEventListener("dragover", onReviewImportDragOver);
  els.reviewsImportDropzone.addEventListener("dragleave", onReviewImportDragLeave);
  els.reviewsImportDropzone.addEventListener("drop", onReviewImportDrop);
  els.reviewsImportDropzone.addEventListener("paste", onReviewImportPaste);
  els.reviewsConfirmImport.addEventListener("click", confirmReviewImport);
  els.reviewsStagingRows.addEventListener("change", onReviewStagingToggle);
  els.reviewsImportRuns.addEventListener("click", onReviewImportRunClick);
  els.reviewsAddProperty.addEventListener("click", createReviewProperty);
  els.reviewsPropertiesBody.addEventListener("click", onReviewPropertyAction);
  els.reviewsSaveSources.addEventListener("click", saveReviewSettings);
  els.reviewsGoogleConnect.addEventListener("click", connectGoogleBusiness);
  els.reviewsGoogleLoadLocations.addEventListener("click", loadGoogleBusinessLocations);
  els.reviewsGoogleSaveMapping.addEventListener("click", saveGoogleBusinessMapping);
  els.reviewsGoogleSync.addEventListener("click", syncGoogleBusinessReviews);
  els.reviewsGoogleMappingsBody?.addEventListener("click", onReviewGoogleMappingAction);
  els.adminUsersBody.addEventListener("click", (event) => {
    const btn = event.target.closest("button[data-action]");
    if (!btn) return;
    const action = clean(btn.dataset.action);
    const userId = clean(btn.dataset.id);
    if (!userId) return;
    if (action === "save-user-profile") {
      saveUserProfile(userId);
      return;
    }
    if (action === "reset-user-password") {
      resetAdminUserPassword(userId);
    }
  });
  els.addProfile.addEventListener("click", createProfile);
  els.profilesBody.addEventListener("click", onProfileAction);
  window.addEventListener("resize", syncMobileNavLayout);
}

async function initAuth() {
  const cfg = window.APP_CONFIG || {};
  const url = clean(cfg.SUPABASE_URL);
  const key = clean(cfg.SUPABASE_ANON_KEY);
  if (!window.supabase || !url || !key) return window.location.replace("/gate.html");
  state.supabase = window.supabase.createClient(url, key);
  const { data, error } = await state.supabase.auth.getSession();
  if (error || !data?.session?.user) return window.location.replace("/gate.html");
  state.user = data.session.user;
  els.authUser.textContent = `Signed in: ${state.user.email || "user"}`;
  state.supabase.auth.onAuthStateChange((_e, session) => {
    if (!session?.user) window.location.replace("/gate.html");
  });
}

async function loadAccess() {
  try {
    const result = await api("/api/access");
    state.access.profile = result.profile || state.access.profile;
    state.access.appFeatures = normalizeFeatureListClient(result.appFeatures, APP_FEATURE_OPTIONS);
    state.access.settingsFeatures = normalizeFeatureListClient(result.settingsFeatures, SETTINGS_FEATURE_OPTIONS);
  } catch (e) {
    state.access = {
      profile: { id: "", name: "Full access (fallback)" },
      appFeatures: [...APP_FEATURE_OPTIONS],
      settingsFeatures: [...SETTINGS_FEATURE_OPTIONS],
    };
    showToast(`Access fallback enabled: ${e.message}`, "info");
  }
}

function normalizeFeatureListClient(list, allowed) {
  const seen = new Set();
  return (Array.isArray(list) ? list : [])
    .map((x) => clean(x).toLowerCase())
    .filter((x) => allowed.includes(x))
    .filter((x) => (seen.has(x) ? false : (seen.add(x), true)));
}

function canApp(feature) {
  return state.access.appFeatures.includes(clean(feature).toLowerCase());
}

function canSettings(feature) {
  return state.access.settingsFeatures.includes(clean(feature).toLowerCase());
}

function canAccessGeneralSettings() {
  return canSettings("general") || canSettings("communications");
}

function isMobileNavLayout() {
  return typeof window !== "undefined" && window.innerWidth <= 768;
}

function setMobileNavOpen(nextOpen) {
  state.mobileNavOpen = !!nextOpen && isMobileNavLayout() && state.currentView !== "settings";
  if (els.appShell) els.appShell.classList.toggle("mobile-nav-open", state.mobileNavOpen);
  if (els.mobileMenuToggle) {
    els.mobileMenuToggle.setAttribute("aria-expanded", state.mobileNavOpen ? "true" : "false");
    els.mobileMenuToggle.setAttribute("aria-label", state.mobileNavOpen ? "Close menu" : "Open menu");
    els.mobileMenuToggle.title = state.mobileNavOpen ? "Close menu" : "Open menu";
  }
}

function toggleMobileNav() {
  setMobileNavOpen(!state.mobileNavOpen);
}

function releaseCommunicationsSearchGuard() {
  if (els.search?.hasAttribute("readonly")) els.search.removeAttribute("readonly");
}

function releaseGuestsIssuerCountryGuard(event) {
  const target = event?.target;
  if (!(target instanceof HTMLElement)) return;
  if (clean(target.dataset?.field) !== "issuerCountry" || clean(target.dataset?.scope) !== "new") return;
  if (target.hasAttribute("readonly")) target.removeAttribute("readonly");
}

function syncMobileNavLayout() {
  if (!isMobileNavLayout()) {
    state.mobileNavOpen = false;
    if (els.appShell) els.appShell.classList.remove("mobile-nav-open");
  }
  setMobileNavOpen(state.mobileNavOpen);
}

function isAdministratorProfile() {
  return clean(state.access?.profile?.name).toLowerCase() === "administrator";
}

function applyInitialRouteFromUrl() {
  try {
    const params = new URLSearchParams(window.location.search);
    const view = clean(params.get("view")).toLowerCase();
    const serviceId = clean(params.get("service") || params.get("request"));
    if (view === "services" && canApp("services")) {
      state.currentView = "services";
    }
    if (view === "guests" && canApp("guests")) {
      state.currentView = "guests";
    }
    if (view === "shopping" && canApp("shopping")) {
      state.currentView = "shopping";
    }
    if (view === "cash" && canApp("cash")) {
      state.currentView = "cash";
    }
    if (view === "hours" && canApp("hours")) {
      state.currentView = "hours";
    }
    if (view === "bakery" && canApp("bakery")) {
      state.currentView = "bakery";
    }
    if (serviceId && canApp("services")) {
      state.pendingServiceDeepLinkId = serviceId;
      state.currentView = "services";
    }
  } catch {}
}

function syncAppRoute() {
  try {
    const url = new URL(window.location.href);
    if (state.currentView === "services") {
      url.searchParams.set("view", "services");
      if (clean(state.serviceSelectedId)) url.searchParams.set("service", clean(state.serviceSelectedId));
      else url.searchParams.delete("service");
    } else if (state.currentView === "cash") {
      url.searchParams.set("view", "cash");
      url.searchParams.delete("service");
    } else if (state.currentView === "guests") {
      url.searchParams.set("view", "guests");
      url.searchParams.delete("service");
    } else if (state.currentView === "hours") {
      url.searchParams.set("view", "hours");
      url.searchParams.delete("service");
    } else {
      url.searchParams.delete("view");
      url.searchParams.delete("service");
    }
    const next = `${url.pathname}${url.search}${url.hash}`;
    window.history.replaceState({}, "", next);
  } catch {}
}

async function setView(view) {
  if (view === "settings" && !state.access.settingsFeatures.length) return showToast("No settings access.", "error");
  if (view === "guests" && !canApp("guests")) return showToast("No guests access.", "error");
  if (view === "lost-found" && !canApp("lost-found")) return showToast("No Lost&Found access.", "error");
  if (view === "reviews" && !canApp("reviews")) return showToast("No reviews access.", "error");
  if (view === "groups" && !canApp("groups")) return showToast("No groups access.", "error");
  if (view === "services" && !canApp("services")) return showToast("No services access.", "error");
  if (view === "cash" && !canApp("cash")) return showToast("No cash control access.", "error");
  if (view === "shopping" && !canApp("shopping")) return showToast("No shopping access.", "error");
  if (view === "hours" && !canApp("hours")) return showToast("No hours register access.", "error");
  if (view === "bakery" && !canApp("bakery")) return showToast("No bakery access.", "error");
  if (view === "laundry" && !canApp("laundry")) return showToast("No laundry access.", "error");
  setMobileNavOpen(false);
  state.currentView = view;
  if (view === "settings") {
    if (canAccessGeneralSettings()) state.settingsSection = "general";
    else if (canSettings("guests")) state.settingsSection = "guests";
    else if (canSettings("cash")) state.settingsSection = "cash";
    else if (canSettings("reviews")) state.settingsSection = "reviews";
    else if (canSettings("groups")) state.settingsSection = "groups";
    else if (canSettings("services")) state.settingsSection = "services";
    else if (canSettings("shopping")) state.settingsSection = "shopping";
    else if (canSettings("hours")) state.settingsSection = "hours";
    else if (canSettings("bakery")) state.settingsSection = "bakery";
    else if (canSettings("laundry")) state.settingsSection = "laundry";
    else if (canSettings("admin-users")) state.settingsSection = "admin-users";
  }
  if (view !== "services") {
    state.serviceSelectedId = "";
    state.pendingServiceDeepLinkId = "";
  }
  if (view !== "shopping" && els.shoppingDetailModal && !els.shoppingDetailModal.hidden) {
    closeShoppingDetailModal();
  }
  if (view !== "bakery" && els.bakeryDetailModal && !els.bakeryDetailModal.hidden) {
    closeBakeryDetailModal();
  }
  if (view !== "cash" && state.cashMoneyModalOpen) {
    closeCashMoneyModal();
  }
  if (view !== "cash" && state.cashItemsModalOpen) {
    closeCashItemsModal();
  }
  syncAppRoute();
  renderLayout();
  renderSettingsSection();
  render();
  await ensureCurrentViewData();
}

async function ensureCurrentViewData() {
  if (state.currentView === "communications") {
    await ensureCommunicationsData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "guests") {
    await ensureGuestsData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "lost-found") {
    await ensureLostFoundData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "reviews") {
    await ensureReviewsData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "groups") {
    await ensureGroupsData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "services") {
    await ensureServicesData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "cash") {
    await ensureCashData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "shopping") {
    await ensureShoppingData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "hours") {
    await ensureHoursData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "bakery") {
    await ensureBakeryData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "laundry") {
    await ensureLaundryData();
    renderSettingsSection();
    render();
    return;
  }
  if (state.currentView === "settings") {
    await ensureSettingsSectionData();
    renderSettingsSection();
    render();
  }
}

function startAutoRefresh() {
  if (state.autoRefreshTimer) return;
  state.autoRefreshTimer = window.setInterval(() => {
    refreshCurrentViewData("timer").catch((error) => console.warn("Auto refresh failed", error));
  }, AUTO_REFRESH_INTERVAL_MS);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) refreshCurrentViewData("focus").catch((error) => console.warn("Auto refresh failed", error));
  });
  window.addEventListener("focus", () => {
    refreshCurrentViewData("focus").catch((error) => console.warn("Auto refresh failed", error));
  });
}

async function refreshCurrentViewData(reason = "timer") {
  if (state.autoRefreshRunning || document.hidden || shouldSkipAutoRefresh()) return;
  const now = Date.now();
  if (reason === "focus" && now - state.lastAutoRefreshAt < AUTO_REFRESH_ON_FOCUS_AFTER_MS) return;
  state.autoRefreshRunning = true;
  try {
    if (state.currentView === "communications" && canApp("communications")) {
      await loadEntries({ silent: true });
      state.communicationsLoaded = true;
      render();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "guests" && canApp("guests")) {
      await loadGuestsData({ silent: true });
      state.guestsLoaded = true;
      renderGuests();
      renderLayout();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "lost-found" && canApp("lost-found")) {
      await loadLostFound({ silent: true });
      state.lostFoundLoaded = true;
      renderLostFound();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "reviews" && canApp("reviews")) {
      await loadReviews({ useFilters: true, silent: true });
      state.reviewsLoaded = true;
      render();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "groups" && canApp("groups")) {
      await loadGroups({ silent: true });
      state.groupsLoaded = true;
      renderGroups();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "services" && canApp("services")) {
      await loadServices({ silent: true });
      state.servicesLoaded = true;
      renderServices();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "cash" && canApp("cash")) {
      await loadCashData({ silent: true });
      state.cashLoaded = true;
      renderCash();
      renderLayout();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "shopping" && canApp("shopping")) {
      await loadShoppingData({ silent: true });
      state.shoppingLoaded = true;
      renderShopping();
      renderLayout();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "hours" && canApp("hours")) {
      await loadHoursData({ silent: true });
      state.hoursLoaded = true;
      renderHours();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "bakery" && canApp("bakery")) {
      await loadBakeryData({ silent: true });
      state.bakeryLoaded = true;
      renderBakery();
      renderLayout();
      state.lastAutoRefreshAt = now;
      return;
    }
    if (state.currentView === "laundry" && canApp("laundry")) {
      await loadLaundryRecords({ silent: true });
      state.laundryLoaded = true;
      renderLaundry();
      state.lastAutoRefreshAt = now;
    }
  } finally {
    state.autoRefreshRunning = false;
  }
}

function shouldSkipAutoRefresh() {
  if (state.currentView === "settings") return true;
  if (state.currentView === "communications" && (state.editingId || hasCommunicationDraft())) return true;
  if (state.currentView === "guests" && (state.guestsEditingId || state.guestsBlacklistEditingId || hasGuestsDraft() || hasGuestsBlacklistDraft())) return true;
  if (state.currentView === "cash" && (state.cashEditingId || hasCashDraft() || state.cashMoneyModalOpen || state.cashItemsModalOpen)) return true;
  if (state.currentView === "lost-found" && (state.lostFoundEditingId || hasLostFoundDraft())) return true;
  if (state.currentView === "groups" && els.groupEditorModal && !els.groupEditorModal.hidden) return true;
  if (state.currentView === "services" && els.serviceEditorModal && !els.serviceEditorModal.hidden) return true;
  if (state.currentView === "shopping" && state.shoppingOpenOrder) return true;
  if (state.currentView === "hours" && (state.hoursEditingId || hasHoursDraft())) return true;
  if (state.currentView === "bakery" && state.bakeryOpenOrder) return true;
  if (state.currentView === "laundry" && els.laundryEditorModal && !els.laundryEditorModal.hidden) return true;
  if (state.reviewQa.loading) return true;
  return false;
}

function hasCommunicationDraft() {
  return !!(clean(state.newDraft.person) || clean(state.newDraft.message));
}

function hasLostFoundDraft() {
  const draft = state.lostFoundDraft || {};
  return !!(
    clean(draft.whoFound) ||
    clean(draft.whoRecorded) ||
    clean(draft.location) ||
    clean(draft.objectDescription) ||
    clean(draft.notes)
  );
}

async function ensureCommunicationsData() {
  if (canApp("communications") && !state.communicationsSettingsLoaded) {
    await loadSettings();
    state.communicationsSettingsLoaded = true;
    renderSettings();
    renderCategoryFilterOptions();
  }
  if (canApp("communications") && !state.communicationsLoaded) {
    await loadEntries();
    state.communicationsLoaded = true;
  }
}

async function ensureGuestsData() {
  if (!canApp("guests") && !canSettings("guests")) return;
  if ((canApp("guests") || canSettings("guests")) && !state.guestsSettingsLoaded) {
    await loadGuestsSettings();
  }
  if (canApp("guests") && !state.guestsLoaded) {
    await loadGuestsData();
    state.guestsLoaded = true;
  }
  renderGuests();
  renderGuestsSettings();
}

async function ensureLostFoundData() {
  if (canApp("lost-found") && !state.lostFoundLoaded) {
    await loadLostFound();
    state.lostFoundLoaded = true;
  }
}

async function ensureReviewsData({ includeImportRuns = false } = {}) {
  if (!canApp("reviews") && !canSettings("reviews")) return;
  if (!state.reviewDateFilterApplied) {
    applyDefaultReviewDateFilter();
    state.reviewDateFilterApplied = true;
  }
  if (!state.reviewPropertiesLoaded) {
    await loadReviewProperties();
    state.reviewPropertiesLoaded = true;
  }
  if (canSettings("reviews") && !state.reviewSettingsLoaded) {
    await loadReviewSettings();
    state.reviewSettingsLoaded = true;
  }
  if (canSettings("reviews") && state.currentView === "settings" && state.settingsSection === "reviews" && state.reviewSettingsScreen === "config" && !state.reviewGoogleLoaded) {
    await loadGoogleBusinessStatus();
    state.reviewGoogleLoaded = true;
  }
  if (canApp("reviews") && !state.reviewsLoaded) {
    await loadReviews();
    state.reviewsLoaded = true;
  }
  if (includeImportRuns && canApp("reviews") && !state.reviewImportRunsLoaded) {
    await loadReviewImportRuns();
    state.reviewImportRunsLoaded = true;
  }
  renderReviewPropertyOptions();
  renderReviewSettings();
}

async function ensureGroupsData() {
  if (!canApp("groups") && !canSettings("groups")) return;
  if (canSettings("groups") && !state.groupSettingsLoaded) {
    await loadGroupSettings();
    state.groupSettingsLoaded = true;
  }
  if (canApp("groups") && !state.groupsLoaded) {
    await loadGroups();
    state.groupsLoaded = true;
  }
  renderGroups();
  renderGroupSettings();
}

async function ensureServicesData() {
  if (!canApp("services") && !canSettings("services")) return;
  if ((canApp("services") || canSettings("services")) && !state.serviceSettingsLoaded) {
    await loadServiceSettings();
    state.serviceSettingsLoaded = true;
  }
  if (canApp("services") && !state.servicesLoaded) {
    await loadServices();
    state.servicesLoaded = true;
  }
  await tryOpenDeepLinkedService();
  renderServices();
  renderServiceSettings();
}

async function ensureShoppingData() {
  if (!canApp("shopping") && !canSettings("shopping")) return;
  if (canSettings("shopping") && !state.shoppingSettingsLoaded) {
    await loadShoppingSettings();
    state.shoppingSettingsLoaded = true;
  }
  if (canApp("shopping") && !state.shoppingLoaded) {
    await loadShoppingData();
    state.shoppingLoaded = true;
  }
  renderShopping();
  renderShoppingSettings();
}

async function ensureCashData() {
  if (!canApp("cash") && !canSettings("cash")) return;
  if (canSettings("cash") && !state.cashSettingsLoaded) {
    await loadCashSettings();
    state.cashSettingsLoaded = true;
  }
  if (canApp("cash") && !state.cashLoaded) {
    await loadCashData();
    state.cashLoaded = true;
  }
  renderCash();
  renderCashSettings();
}

async function ensureHoursData() {
  if (!canApp("hours") && !canSettings("hours")) return;
  if (canSettings("hours") && !state.hoursSettingsLoaded) {
    await loadHoursSettings();
    state.hoursSettingsLoaded = true;
  }
  if (canApp("hours") && !state.hoursLoaded) {
    await loadHoursData();
    state.hoursLoaded = true;
  }
  renderHours();
  renderHoursSettings();
}

async function ensureBakeryData() {
  if (!canApp("bakery") && !canSettings("bakery")) return;
  if (canSettings("bakery") && !state.bakerySettingsLoaded) {
    await loadBakerySettings();
    state.bakerySettingsLoaded = true;
  }
  if (canApp("bakery") && !state.bakeryLoaded) {
    await loadBakeryData();
    state.bakeryLoaded = true;
  }
  renderBakery();
  renderBakerySettings();
}

async function ensureLaundryData() {
  if (!canApp("laundry") && !canSettings("laundry")) return;
  if ((canApp("laundry") || canSettings("laundry")) && !state.laundrySettingsLoaded) {
    await loadLaundrySettings();
    state.laundrySettingsLoaded = true;
  }
  if (canApp("laundry") && !state.laundryLoaded) {
    await loadLaundryRecords();
    state.laundryLoaded = true;
  }
  renderLaundry();
  renderLaundrySettings();
}

async function ensureSettingsSectionData() {
  if (state.settingsSection === "general") {
    await ensureCommunicationsData();
    return;
  }
  if (state.settingsSection === "guests") {
    await ensureGuestsData();
    return;
  }
  if (state.settingsSection === "communications") {
    await ensureCommunicationsData();
    return;
  }
  if (state.settingsSection === "reviews") {
    await ensureReviewsData({ includeImportRuns: state.reviewSettingsScreen === "import" });
    return;
  }
  if (state.settingsSection === "groups") {
    await ensureGroupsData();
    return;
  }
  if (state.settingsSection === "services") {
    await ensureServicesData();
    return;
  }
  if (state.settingsSection === "cash") {
    await ensureCashData();
    return;
  }
  if (state.settingsSection === "shopping") {
    await ensureShoppingData();
    return;
  }
  if (state.settingsSection === "hours") {
    await ensureHoursData();
    return;
  }
  if (state.settingsSection === "bakery") {
    await ensureBakeryData();
    return;
  }
  if (state.settingsSection === "laundry") {
    await ensureLaundryData();
    return;
  }
  if (state.settingsSection === "admin-users") await ensureAdminUsersData();
}

function renderLayout() {
  const comm = state.currentView === "communications";
  const guests = state.currentView === "guests";
  const cash = state.currentView === "cash";
  const lostFound = state.currentView === "lost-found";
  const reviews = state.currentView === "reviews";
  const groups = state.currentView === "groups";
  const services = state.currentView === "services";
  const shopping = state.currentView === "shopping";
  const hours = state.currentView === "hours";
  const bakery = state.currentView === "bakery";
  const laundry = state.currentView === "laundry";
  const settingsMode = state.currentView === "settings";
  const canComm = canApp("communications");
  const canGuests = canApp("guests");
  const canCash = canApp("cash");
  const canLostFound = canApp("lost-found");
  const canReviews = canApp("reviews");
  const canGroups = canApp("groups");
  const canServices = canApp("services");
  const canShopping = canApp("shopping");
  const canHours = canApp("hours");
  const canBakery = canApp("bakery");
  const canLaundry = canApp("laundry");
  if (els.sidebarReviewSummaryCard) els.sidebarReviewSummaryCard.hidden = !canComm;

  els.appShell.classList.toggle("settings-mode", settingsMode);
  els.navCommunications.classList.toggle("active", comm);
  els.navGuests?.classList.toggle("active", guests);
  els.navCash?.classList.toggle("active", cash);
  els.navLostFound.classList.toggle("active", lostFound);
  els.navReviews.classList.toggle("active", reviews);
  els.navGroups.classList.toggle("active", groups);
  els.navServices.classList.toggle("active", services);
  els.navShopping.classList.toggle("active", shopping);
  els.navHours.classList.toggle("active", hours);
  els.navBakery.classList.toggle("active", bakery);
  els.navLaundry.classList.toggle("active", laundry);
  els.navCommunications.hidden = !canComm;
  if (els.navGuests) els.navGuests.hidden = !canGuests;
  if (els.navCash) els.navCash.hidden = !canCash;
  els.navLostFound.hidden = !canLostFound;
  els.navReviews.hidden = !canReviews;
  els.navGroups.hidden = !canGroups;
  els.navServices.hidden = !canServices;
  els.navShopping.hidden = !canShopping;
  els.navHours.hidden = !canHours;
  els.navBakery.hidden = !canBakery;
  els.navLaundry.hidden = !canLaundry;
  els.navGuests?.classList.toggle("has-alert", shouldShowGuestsAlertClient());
  els.navCash?.classList.toggle("has-alert", shouldShowCashAlert());
  els.navShopping.classList.toggle("has-alert", shouldShowShoppingAlert());
  els.navHours.classList.toggle("has-alert", shouldShowHoursAlert());
  els.navBakery.classList.toggle("has-alert", shouldShowBakeryAlert());
  els.navLaundry.classList.toggle("has-alert", shouldShowLaundryAlert());
  els.openSettings.hidden = !state.access.settingsFeatures.length;
  els.leftNav.hidden = settingsMode;
  els.topbar.hidden = false;
  if (els.mobileMenuToggle) els.mobileMenuToggle.hidden = settingsMode || !isMobileNavLayout();
  els.viewCommunications.hidden = !comm;
  if (els.viewGuests) els.viewGuests.hidden = !guests;
  if (els.viewCash) els.viewCash.hidden = !cash;
  els.viewLostFound.hidden = !lostFound;
  els.viewReviews.hidden = !reviews;
  els.viewGroups.hidden = !groups;
  els.viewServices.hidden = !services;
  els.viewShopping.hidden = !shopping;
  els.viewHours.hidden = !hours;
  els.viewBakery.hidden = !bakery;
  els.viewLaundry.hidden = !laundry;
  els.viewSettings.hidden = !settingsMode;
  els.settingsMenuGeneral.hidden = !canAccessGeneralSettings();
  els.settingsMenuCommunications.hidden = !canSettings("communications");
  if (els.settingsMenuGuests) els.settingsMenuGuests.hidden = !canSettings("guests");
  if (els.settingsMenuCash) els.settingsMenuCash.hidden = !canSettings("cash");
  els.settingsMenuReviews.hidden = !canSettings("reviews");
  els.settingsMenuGroups.hidden = !canSettings("groups");
  els.settingsMenuServices.hidden = !canSettings("services");
  els.settingsMenuShopping.hidden = !canSettings("shopping");
  els.settingsMenuHours.hidden = !canSettings("hours");
  els.settingsMenuBakery.hidden = !canSettings("bakery");
  els.settingsMenuLaundry.hidden = !canSettings("laundry");
  els.settingsMenuAdminUsers.hidden = !canSettings("admin-users");
  els.settingsMenuGeneral.classList.toggle("active", state.settingsSection === "general");
  els.settingsMenuCommunications.classList.toggle("active", state.settingsSection === "communications");
  els.settingsMenuGuests?.classList.toggle("active", state.settingsSection === "guests");
  els.settingsMenuCash?.classList.toggle("active", state.settingsSection === "cash");
  els.settingsMenuReviews.classList.toggle("active", state.settingsSection === "reviews");
  els.settingsMenuGroups.classList.toggle("active", state.settingsSection === "groups");
  els.settingsMenuServices.classList.toggle("active", state.settingsSection === "services");
  els.settingsMenuShopping.classList.toggle("active", state.settingsSection === "shopping");
  els.settingsMenuHours.classList.toggle("active", state.settingsSection === "hours");
  els.settingsMenuBakery.classList.toggle("active", state.settingsSection === "bakery");
  els.settingsMenuLaundry.classList.toggle("active", state.settingsSection === "laundry");
  els.settingsMenuAdminUsers.classList.toggle("active", state.settingsSection === "admin-users");
  renderSidebarReviewSummary();
  syncMobileNavLayout();
}

async function setSettingsSection(section) {
  if (section === "general" && !canAccessGeneralSettings()) return;
  if (section === "guests" && !canSettings("guests")) return;
  if (section === "communications" && !canSettings("communications")) return;
  if (section === "cash" && !canSettings("cash")) return;
  if (section === "reviews" && !canSettings("reviews")) return;
  if (section === "groups" && !canSettings("groups")) return;
  if (section === "services" && !canSettings("services")) return;
  if (section === "shopping" && !canSettings("shopping")) return;
  if (section === "hours" && !canSettings("hours")) return;
  if (section === "bakery" && !canSettings("bakery")) return;
  if (section === "laundry" && !canSettings("laundry")) return;
  if (section === "admin-users" && !canSettings("admin-users")) return;
  setMobileNavOpen(false);
  if (section === "guests") state.guestsSettingsLoaded = false;
  state.settingsSection = section === "admin-users"
    ? "admin-users"
    : section === "guests"
      ? "guests"
    : section === "cash"
      ? "cash"
    : section === "shopping"
      ? "shopping"
    : section === "hours"
      ? "hours"
    : section === "bakery"
      ? "bakery"
    : section === "laundry"
      ? "laundry"
    : section === "services"
      ? "services"
      : section === "groups"
        ? "groups"
        : section === "reviews"
          ? "reviews"
          : section === "communications"
            ? "communications"
            : "general";
  renderLayout();
  renderSettingsSection();
  await ensureSettingsSectionData();
  renderSettingsSection();
  render();
}

async function ensureAdminUsersData() {
  if (state.adminUsersLoaded) return;
  try {
    await loadProfiles();
    renderProfiles();
  } catch (e) {
    setProfilesStatus(`Failed to load profiles: ${e.message}`);
  }
  try {
    await loadAdminUsers();
  } catch (e) {
    setAdminUsersStatus(`Failed to load users: ${e.message}`);
  }
  renderAdminUsers();
}

function renderSettingsSection() {
  const isGeneral = state.settingsSection === "general" && canAccessGeneralSettings();
  const isGuests = state.settingsSection === "guests" && canSettings("guests");
  const isComm = state.settingsSection === "communications" && canSettings("communications");
  const isCash = state.settingsSection === "cash" && canSettings("cash");
  const isReviews = state.settingsSection === "reviews" && canSettings("reviews");
  const isGroups = state.settingsSection === "groups" && canSettings("groups");
  const isServices = state.settingsSection === "services" && canSettings("services");
  const isShopping = state.settingsSection === "shopping" && canSettings("shopping");
  const isHours = state.settingsSection === "hours" && canSettings("hours");
  const isBakery = state.settingsSection === "bakery" && canSettings("bakery");
  const isLaundry = state.settingsSection === "laundry" && canSettings("laundry");
  const isAdmin = state.settingsSection === "admin-users" && canSettings("admin-users");
  els.settingsViewGeneral.hidden = !isGeneral;
  if (els.settingsViewGuests) els.settingsViewGuests.hidden = !isGuests;
  els.settingsViewCommunications.hidden = !isComm;
  if (els.settingsViewCash) els.settingsViewCash.hidden = !isCash;
  els.settingsViewReviews.hidden = !isReviews;
  els.settingsViewGroups.hidden = !isGroups;
  els.settingsViewServices.hidden = !isServices;
  els.settingsViewShopping.hidden = !isShopping;
  els.settingsViewHours.hidden = !isHours;
  els.settingsViewBakery.hidden = !isBakery;
  els.settingsViewLaundry.hidden = !isLaundry;
  els.settingsViewAdminUsers.hidden = !isAdmin;
  if (isReviews) setReviewSettingsScreen(state.reviewSettingsScreen, false);
}

async function signOut() {
  await state.supabase.auth.signOut();
  window.location.replace("/gate.html");
}

async function api(path, options = {}) {
  const { data } = await state.supabase.auth.getSession();
  const token = data?.session?.access_token;
  const headers = {
    Authorization: `Bearer ${token}`,
  };
  if (options.body !== undefined) headers["Content-Type"] = "application/json";
  const response = await fetch(path, {
    method: options.method || "GET",
    headers,
    body: options.body === undefined ? undefined : JSON.stringify(options.body),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || `Request failed (${response.status})`);
  return payload;
}

async function loadSettings() {
  try {
    const result = await api("/api/settings");
    state.settings = sanitizeSettings(result.settings);
    setSettingsStatus("Settings loaded.");
    setGeneralSettingsStatus("Settings loaded.");
  } catch (e) {
    state.settings = clone(DEFAULT_SETTINGS);
    setSettingsStatus(`Using defaults (${e.message}).`);
    setGeneralSettingsStatus(`Using defaults (${e.message}).`);
  }
}

async function saveSettings() {
  state.settings = sanitizeSettings(state.settings);
  normalizeDraftsToSettings();
  renderCategoryFilterOptions();
  renderSettings();
  render();
  try {
    await api("/api/settings", { method: "PUT", body: { settings: state.settings } });
    setSettingsStatus("Settings saved.");
    setGeneralSettingsStatus("Settings saved.");
    showToast("Settings saved.", "success");
  } catch (e) {
    setSettingsStatus(`Save failed: ${e.message}`);
    setGeneralSettingsStatus(`Save failed: ${e.message}`);
    showToast(`Settings save failed: ${e.message}`, "error");
  }
}

async function loadAdminUsers(forceRefresh = false) {
  if (state.adminUsersLoaded && !forceRefresh) return;
  setAdminUsersStatus("Loading users...");
  const result = await api("/api/admin-users");
  state.adminUsers = Array.isArray(result.users) ? result.users : [];
  state.adminUsersLoaded = true;
  renderAdminUsers();
  setAdminUsersStatus(`Loaded ${state.adminUsers.length} user${state.adminUsers.length === 1 ? "" : "s"}.`);
}

async function loadProfiles(forceRefresh = false) {
  if (state.profilesLoaded && !forceRefresh) return;
  setProfilesStatus("Loading profiles...");
  const result = await api("/api/profiles");
  state.profiles = (Array.isArray(result.profiles) ? result.profiles : []).map((p) => ({
    id: clean(p.id),
    name: clean(p.name),
    appFeatures: normalizeFeatureListClient(p.app_features || p.appFeatures, APP_FEATURE_OPTIONS),
    settingsFeatures: normalizeFeatureListClient(p.settings_features || p.settingsFeatures, SETTINGS_FEATURE_OPTIONS),
  }));
  state.profilesLoaded = true;
  renderProfiles();
  renderAdminUsers();
  setProfilesStatus(`Loaded ${state.profiles.length} profile${state.profiles.length === 1 ? "" : "s"}.`);
}

async function createAdminUser() {
  const email = clean(els.adminUserEmail.value).toLowerCase();
  const password = String(els.adminUserPassword.value || "");
  const profileId = clean(els.adminUserProfile.value);
  if (!email || !email.includes("@")) return setAdminUsersStatus("Please provide a valid email.");
  if (password.length < 8) return setAdminUsersStatus("Password must have at least 8 characters.");

  els.adminCreateUser.disabled = true;
  try {
    await api("/api/admin-users", {
      method: "POST",
      body: { email, password, profileId },
    });
    els.adminUserEmail.value = "";
    els.adminUserPassword.value = "";
    els.adminUserProfile.value = "";
    state.adminUsersLoaded = false;
    await loadAdminUsers(true);
    setAdminUsersStatus("User created successfully.");
    showToast("User created successfully.", "success");
  } catch (e) {
    setAdminUsersStatus(`Create failed: ${e.message}`);
    showToast(`Create failed: ${e.message}`, "error");
  } finally {
    els.adminCreateUser.disabled = false;
  }
}

function renderAdminUsers() {
  els.adminUsersBody.innerHTML = "";
  const profileOptions = `<option value="">(no profile)</option>${state.profiles
    .map((p) => `<option value="${escape(p.id)}">${escape(p.name)}</option>`)
    .join("")}`;

  els.adminUserProfile.innerHTML = profileOptions;
  if (state.adminUsers.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td colspan="6" class="empty">No users found.</td>';
    els.adminUsersBody.appendChild(tr);
    return;
  }

  state.adminUsers.forEach((user) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(user.email || "(no email)")}</td>
      <td>${escape(formatDateTimeShort(user.createdAt))}</td>
      <td>${escape(formatDateTimeShort(user.lastSignInAt))}</td>
      <td>${escape(user.emailConfirmedAt ? "Yes" : "No")}</td>
      <td>
        <select data-user-profile="${escape(user.id)}">${profileOptions}</select>
      </td>
      <td class="row-actions admin-users-action">
        <button type="button" class="ghost" data-action="save-user-profile" data-id="${escape(user.id)}">Save</button>
        <button type="button" class="ghost" data-action="reset-user-password" data-id="${escape(user.id)}">Reset Password</button>
      </td>`;
    const select = tr.querySelector("select");
    if (select) select.value = clean(user.profileId);
    els.adminUsersBody.appendChild(tr);
  });
}

function renderProfiles() {
  if (els.profilesHead) {
    els.profilesHead.innerHTML = state.profiles.length
      ? `<tr><th>Feature</th>${state.profiles
          .map((profile) => `<th>${escape(profile.name || "Profile")}</th>`)
          .join("")}</tr>`
      : '<tr><th>Feature</th><th>Profiles</th></tr>';
  }
  els.profilesBody.innerHTML = "";
  if (state.profiles.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<th scope="row">Profiles</th><td class="empty">No profiles yet.</td>';
    els.profilesBody.appendChild(tr);
    return;
  }
  PROFILE_MATRIX_ROWS.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<th scope="row">${escape(row.label)}</th>${state.profiles
      .map((profile) => {
        if (row.kind === "meta") {
          return `<td><input data-profile-name="${escape(profile.id)}" value="${escape(profile.name)}" /></td>`;
        }
        if (row.kind === "action") {
          return `<td class="row-actions center-cell profile-matrix-action">
            <button type="button" class="ghost" data-action="save-profile" data-id="${escape(profile.id)}">Save</button>
            <button type="button" class="danger" data-action="delete-profile" data-id="${escape(profile.id)}">Delete</button>
          </td>`;
        }
        const hasFeature = row.kind === "app"
          ? profile.appFeatures.includes(row.key)
          : profile.settingsFeatures.includes(row.key);
        return `<td class="center-cell"><input type="checkbox" data-profile-${row.kind}-${escape(row.key)}="${escape(profile.id)}" ${hasFeature ? "checked" : ""} /></td>`;
      })
      .join("")}`;
    els.profilesBody.appendChild(tr);
  });
}

async function createProfile() {
  try {
    const created = await api("/api/profiles", {
      method: "POST",
      body: {
        name: `Profile ${state.profiles.length + 1}`,
        appFeatures: [...APP_FEATURE_OPTIONS],
        settingsFeatures: [],
      },
    });
    if (created.profile) {
      state.profilesLoaded = false;
      await loadProfiles(true);
      setProfilesStatus("Profile created.");
      showToast("Profile created.", "success");
    }
  } catch (e) {
    setProfilesStatus(`Create profile failed: ${e.message}`);
    showToast(`Create profile failed: ${e.message}`, "error");
  }
}

async function onProfileAction(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) return;
  const action = clean(button.dataset.action);
  const id = clean(button.dataset.id);
  if (!id) return;

  if (action === "save-profile") {
    const payload = collectProfilePayload(id);
    if (!payload) return;
    try {
      await api(`/api/profiles?id=${encodeURIComponent(id)}`, { method: "PUT", body: payload });
      state.profilesLoaded = false;
      await loadProfiles(true);
      setProfilesStatus("Profile saved.");
      showToast("Profile saved.", "success");
    } catch (e) {
      setProfilesStatus(`Save profile failed: ${e.message}`);
      showToast(`Save profile failed: ${e.message}`, "error");
    }
    return;
  }

  if (action === "delete-profile") {
    if (!window.confirm("Delete this profile?")) return;
    try {
      await api(`/api/profiles?id=${encodeURIComponent(id)}`, { method: "DELETE" });
      state.profilesLoaded = false;
      await loadProfiles(true);
      state.adminUsersLoaded = false;
      await loadAdminUsers(true);
      setProfilesStatus("Profile deleted.");
      showToast("Profile deleted.", "success");
    } catch (e) {
      setProfilesStatus(`Delete profile failed: ${e.message}`);
      showToast(`Delete profile failed: ${e.message}`, "error");
    }
  }
}

function collectProfilePayload(id) {
  const name = clean(els.profilesBody.querySelector(`[data-profile-name="${id}"]`)?.value);
  if (!name) {
    setProfilesStatus("Profile name is required.");
    return null;
  }
  const appFeatures = [];
  if (els.profilesBody.querySelector(`[data-profile-app-communications="${id}"]`)?.checked) appFeatures.push("communications");
  if (els.profilesBody.querySelector(`[data-profile-app-guests="${id}"]`)?.checked) appFeatures.push("guests");
  if (els.profilesBody.querySelector(`[data-profile-app-cash="${id}"]`)?.checked) appFeatures.push("cash");
  if (els.profilesBody.querySelector(`[data-profile-app-lost-found="${id}"]`)?.checked) appFeatures.push("lost-found");
  if (els.profilesBody.querySelector(`[data-profile-app-reviews="${id}"]`)?.checked) appFeatures.push("reviews");
  if (els.profilesBody.querySelector(`[data-profile-app-groups="${id}"]`)?.checked) appFeatures.push("groups");
  if (els.profilesBody.querySelector(`[data-profile-app-services="${id}"]`)?.checked) appFeatures.push("services");
  if (els.profilesBody.querySelector(`[data-profile-app-shopping="${id}"]`)?.checked) appFeatures.push("shopping");
  if (els.profilesBody.querySelector(`[data-profile-app-hours="${id}"]`)?.checked) appFeatures.push("hours");
  if (els.profilesBody.querySelector(`[data-profile-app-bakery="${id}"]`)?.checked) appFeatures.push("bakery");
  if (els.profilesBody.querySelector(`[data-profile-app-laundry="${id}"]`)?.checked) appFeatures.push("laundry");
  const settingsFeatures = [];
  if (els.profilesBody.querySelector(`[data-profile-settings-general="${id}"]`)?.checked) settingsFeatures.push("general");
  if (els.profilesBody.querySelector(`[data-profile-settings-communications="${id}"]`)?.checked) settingsFeatures.push("communications");
  if (els.profilesBody.querySelector(`[data-profile-settings-guests="${id}"]`)?.checked) settingsFeatures.push("guests");
  if (els.profilesBody.querySelector(`[data-profile-settings-cash="${id}"]`)?.checked) settingsFeatures.push("cash");
  if (els.profilesBody.querySelector(`[data-profile-settings-reviews="${id}"]`)?.checked) settingsFeatures.push("reviews");
  if (els.profilesBody.querySelector(`[data-profile-settings-groups="${id}"]`)?.checked) settingsFeatures.push("groups");
  if (els.profilesBody.querySelector(`[data-profile-settings-services="${id}"]`)?.checked) settingsFeatures.push("services");
  if (els.profilesBody.querySelector(`[data-profile-settings-shopping="${id}"]`)?.checked) settingsFeatures.push("shopping");
  if (els.profilesBody.querySelector(`[data-profile-settings-hours="${id}"]`)?.checked) settingsFeatures.push("hours");
  if (els.profilesBody.querySelector(`[data-profile-settings-bakery="${id}"]`)?.checked) settingsFeatures.push("bakery");
  if (els.profilesBody.querySelector(`[data-profile-settings-laundry="${id}"]`)?.checked) settingsFeatures.push("laundry");
  if (els.profilesBody.querySelector(`[data-profile-settings-admin-users="${id}"]`)?.checked) settingsFeatures.push("admin-users");
  return { name, appFeatures, settingsFeatures };
}

async function saveUserProfile(userId) {
  const select = els.adminUsersBody.querySelector(`select[data-user-profile="${userId}"]`);
  const profileId = clean(select?.value);
  try {
    await api("/api/admin-users", { method: "PATCH", body: { userId, profileId } });
    const row = state.adminUsers.find((x) => x.id === userId);
    if (row) row.profileId = profileId;
    setAdminUsersStatus("User profile updated.");
    showToast("User profile updated.", "success");
  } catch (e) {
    setAdminUsersStatus(`Profile update failed: ${e.message}`);
    showToast(`Profile update failed: ${e.message}`, "error");
  }
}

async function resetAdminUserPassword(userId) {
  const user = state.adminUsers.find((item) => item.id === userId);
  const email = clean(user?.email) || "this user";
  const password = String(window.prompt(`Enter a new password for ${email}:`, "") || "").trim();
  if (!password) return;
  if (password.length < 8) {
    setAdminUsersStatus("Password must have at least 8 characters.");
    showToast("Password must have at least 8 characters.", "error");
    return;
  }
  const confirmation = String(window.prompt(`Confirm the new password for ${email}:`, "") || "").trim();
  if (password !== confirmation) {
    setAdminUsersStatus("The passwords do not match.");
    showToast("The passwords do not match.", "error");
    return;
  }
  try {
    await api("/api/admin-users", { method: "PATCH", body: { userId, password } });
    setAdminUsersStatus(`Password updated for ${email}.`);
    showToast(`Password updated for ${email}.`, "success");
  } catch (e) {
    setAdminUsersStatus(`Password reset failed: ${e.message}`);
    showToast(`Password reset failed: ${e.message}`, "error");
  }
}

function emptyGroupDraft() {
  return {
    id: "",
    reservationNumber: "",
    creationDate: "",
    name: "",
    email: "",
    checkIn: "",
    checkOut: "",
    guests: "",
    roomItems: [],
    totalValue: 0,
    optionDate: "",
    status: "Proposal",
    observation: "",
    language: "en",
    audit: [],
  };
}

async function loadGroupSettings({ silent = false } = {}) {
  try {
    const result = await api("/api/group-settings");
    state.groupSettings = sanitizeGroupSettings(result.settings);
    state.groupSettingsLoaded = true;
    renderGroupSettings();
  } catch (e) {
    state.groupSettings = clone(DEFAULT_GROUP_SETTINGS);
    if (!silent) setGroupsSettingsStatus(`Using default group settings (${e.message}).`);
  }
}

async function loadGroups({ silent = false } = {}) {
  try {
    const result = await api("/api/groups");
    state.groups = (Array.isArray(result.rows) ? result.rows : []).map(mapGroupRow);
    if (!silent) setGroupsStatus(`Loaded ${state.groups.length} proposal${state.groups.length === 1 ? "" : "s"}.`);
  } catch (e) {
    state.groups = [];
    setGroupsStatus(`Failed to load groups: ${e.message}`);
  }
}

function mapGroupRow(row) {
  const metadata = groupMetadata(row.guest_groups);
  return {
    id: clean(row.id),
    reservationNumber: clean(row.reservation_number),
    creationDate: clean(row.creation_date || row.created_at),
    name: clean(row.name),
    email: clean(row.email),
    checkIn: clean(row.check_in),
    checkOut: clean(row.check_out),
    guests: Number(row.guests || 0),
    roomItems: normalizeGroupRoomItems(row.room_items),
    totalValue: normalizeNumber(row.total_value),
    optionDate: clean(row.option_date),
    status: normalizeGroupStatus(row.status),
    observation: clean(metadata.observation),
    language: normalizeProposalLanguage(metadata.language),
    audit: normalizeGroupAudit(metadata.audit),
  };
}

function groupProposalCountByEmail(email) {
  const needle = clean(email).toLowerCase();
  if (!needle) return 0;
  return state.groups.filter((row) => clean(row.email).toLowerCase() === needle).length;
}

function renderGroupEmailProposalHint() {
  if (!els.groupEmailProposalsHint) return;
  const count = groupProposalCountByEmail(state.groupDraft.email);
  if (count > 1) {
    els.groupEmailProposalsHint.hidden = false;
    els.groupEmailProposalsHint.textContent = `${count} proposals found for this email`;
  } else {
    els.groupEmailProposalsHint.hidden = true;
    els.groupEmailProposalsHint.textContent = "";
  }
}

function setGroupsScreen(screen) {
  state.groupsScreen = screen === "resume" ? "resume" : "list";
  renderGroups();
}

function renderGroupsScreenTabs() {
  const isResume = state.groupsScreen === "resume";
  els.groupsTabList?.classList.toggle("active-tab", !isResume);
  els.groupsTabList?.classList.toggle("ghost", isResume);
  els.groupsTabResume?.classList.toggle("active-tab", isResume);
  els.groupsTabResume?.classList.toggle("ghost", !isResume);
  if (els.groupsPanelList) els.groupsPanelList.hidden = isResume;
  if (els.groupsPanelResume) els.groupsPanelResume.hidden = !isResume;
}

function formatGroupMonthLabel(monthKey) {
  const raw = clean(monthKey);
  if (!/^\d{4}-\d{2}$/.test(raw)) return raw || "-";
  const dt = new Date(`${raw}-01T00:00:00`);
  if (Number.isNaN(dt.getTime())) return raw;
  return new Intl.DateTimeFormat("en-GB", { month: "short", year: "numeric", timeZone: "Europe/Lisbon" }).format(dt);
}

function getGroupResumeRows() {
  const mode = state.groupResumeMonthMode === "checkin" ? "checkin" : "created";
  const buckets = new Map();
  getFilteredGroups().forEach((row) => {
    const sourceDate = mode === "checkin" ? clean(row.checkIn) : clean(row.creationDate).slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sourceDate)) return;
    const monthKey = sourceDate.slice(0, 7);
    const current = buckets.get(monthKey) || {
      monthKey,
      totalProposals: 0,
      totalGuests: 0,
      totalAmount: 0,
      acceptedProposals: 0,
      acceptedGuests: 0,
      acceptedAmount: 0,
    };
    const guests = Math.max(0, Number(row.guests || 0));
    const amount = Math.max(0, Number(row.totalValue || 0));
    const isAccepted = clean(row.status) === "Accepted";
    current.totalProposals += 1;
    current.totalGuests += guests;
    current.totalAmount += amount;
    if (isAccepted) {
      current.acceptedProposals += 1;
      current.acceptedGuests += guests;
      current.acceptedAmount += amount;
    }
    buckets.set(monthKey, current);
  });
  return [...buckets.values()].sort((a, b) => clean(b.monthKey).localeCompare(clean(a.monthKey)));
}

function renderGroupsResume() {
  if (!els.groupsResumeBody || !els.groupsResumeCount) return;
  if (els.groupsResumeMonthMode) els.groupsResumeMonthMode.value = state.groupResumeMonthMode === "checkin" ? "checkin" : "created";
  const rows = getGroupResumeRows();
  els.groupsResumeCount.textContent = `${rows.length} month${rows.length === 1 ? "" : "s"}`;
  els.groupsResumeBody.innerHTML = "";
  if (!rows.length) {
    els.groupsResumeBody.innerHTML = '<tr><td colspan="7" class="empty">No group proposals found.</td></tr>';
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(formatGroupMonthLabel(row.monthKey))}</td>
      <td>${escape(String(row.totalProposals))}</td>
      <td>${escape(String(row.totalGuests))}</td>
      <td>${escape(formatMoney(row.totalAmount))}</td>
      <td>${escape(String(row.acceptedProposals))}</td>
      <td>${escape(String(row.acceptedGuests))}</td>
      <td>${escape(formatMoney(row.acceptedAmount))}</td>`;
    els.groupsResumeBody.appendChild(tr);
  });
}

function renderGroups() {
  if (!els.groupsRows || !canApp("groups")) return;
  renderGroupsScreenTabs();
  if (!els.groupEditorModal.hidden) renderGroupDraft();
  if (els.groupsFilterCreatedFrom) els.groupsFilterCreatedFrom.value = clean(state.groupFilters.createdFrom);
  if (els.groupsFilterCreatedTo) els.groupsFilterCreatedTo.value = clean(state.groupFilters.createdTo);
  if (els.groupsFilterDateFrom) els.groupsFilterDateFrom.value = clean(state.groupFilters.dateFrom);
  if (els.groupsFilterDateTo) els.groupsFilterDateTo.value = clean(state.groupFilters.dateTo);
  if (els.groupsFilterSearch) els.groupsFilterSearch.value = clean(state.groupFilters.search);
  const rows = getFilteredGroups();
  updateGroupSortIndicators();
  els.groupsCount.textContent = `${rows.length} proposal${rows.length === 1 ? "" : "s"}`;
  if (state.groupsScreen === "resume") {
    renderGroupsResume();
    return;
  }
  els.groupsRows.innerHTML = "";
  if (els.groupsMobileCards) els.groupsMobileCards.innerHTML = "";
  if (rows.length === 0) {
    els.groupsRows.innerHTML = '<tr><td colspan="9" class="empty">No group proposals found.</td></tr>';
    if (els.groupsMobileCards) {
      els.groupsMobileCards.innerHTML = '<div class="services-mobile-empty">No group proposals found.</div>';
    }
    return;
  }
  rows.forEach((row) => {
    const roomTypeSummary = groupRoomTypeSummary(row.roomItems);
    const roomSummary = groupRoomSelectionSummary(row.roomItems);
    const tr = document.createElement("tr");
    tr.dataset.groupId = row.id;
    const statusClass = row.status === "Accepted" ? " accepted-row" : row.status === "Refused" ? " refused-row" : "";
    tr.className = `clickable-row${statusClass}${row.id === state.groupSelectedId ? " selected-row" : ""}`;
    tr.innerHTML = `<td>${escape(formatDateOnly(row.creationDate))}</td>
      <td>${escape(row.name)}</td>
      <td class="group-dates-cell"><span>${escape(formatGroupDateDisplay(row.checkIn))} - ${escape(formatGroupDateDisplay(row.checkOut))}</span><small>(${escape(String(dateDiffDays(row.checkIn, row.checkOut)))} nights)</small></td>
      <td>${escape(String(row.guests || 0))}</td>
      <td class="group-total-cell"><strong>${escape(formatMoney(row.totalValue))}</strong><small>(${escape(groupDepositText(row.totalValue))})</small></td>
      <td class="compact-summary-cell room-types-cell" title="${escape(roomTypeSummary)}">${groupRoomTypeSummaryHtml(row.roomItems)}</td>
      <td class="compact-summary-cell" title="${escape(roomSummary)}">${escape(roomSummary || "-")}</td>
      <td>${escape(row.optionDate || "-")}</td>
      <td>${escape(row.reservationNumber || "-")}</td>`;
    els.groupsRows.appendChild(tr);
    if (els.groupsMobileCards) {
      els.groupsMobileCards.appendChild(buildGroupMobileCard(row));
    }
  });
}

function buildGroupMobileCard(row) {
  const roomTypeSummary = groupRoomTypeSummary(row.roomItems);
  const roomSummary = groupRoomSelectionSummary(row.roomItems);
  const card = document.createElement("article");
  const statusClass = row.status === "Accepted" ? " accepted-row" : row.status === "Refused" ? " refused-row" : "";
  card.className = `group-mobile-card${statusClass}${row.id === state.groupSelectedId ? " selected-card" : ""}`;
  card.dataset.groupId = row.id;
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(row.name)}</div>
        <div class="communication-mobile-meta">Created: ${escape(formatDateOnly(row.creationDate))}</div>
      </div>
      <div class="group-mobile-total">
        <strong>${escape(formatMoney(row.totalValue))}</strong>
        <small>${escape(groupDepositText(row.totalValue))}</small>
      </div>
    </div>
    <div class="communication-mobile-grid">
      <div class="communication-mobile-field">
        <small>Dates</small>
        <div class="communication-mobile-message">${escape(formatGroupDateDisplay(row.checkIn))} - ${escape(formatGroupDateDisplay(row.checkOut))}<br><small>${escape(String(dateDiffDays(row.checkIn, row.checkOut)))} nights</small></div>
      </div>
      <div class="communication-mobile-field">
        <small>Guests</small>
        <div class="communication-mobile-message">${escape(String(row.guests || 0))}</div>
      </div>
      <div class="communication-mobile-field">
        <small>Option</small>
        <div class="communication-mobile-message">${escape(row.optionDate || "-")}</div>
      </div>
      <div class="communication-mobile-field">
        <small>Reservation</small>
        <div class="communication-mobile-message">${escape(row.reservationNumber || "-")}</div>
      </div>
      <div class="communication-mobile-field communication-mobile-field-full">
        <small>Room Types</small>
        <div class="communication-mobile-message">${escape(roomTypeSummary || "-")}</div>
      </div>
      <div class="communication-mobile-field communication-mobile-field-full">
        <small>Rooms</small>
        <div class="communication-mobile-message">${escape(roomSummary || "-")}</div>
      </div>
    </div>`;
  return card;
}

function getFilteredGroups() {
  const today = formatDate(new Date());
  const createdFrom = clean(state.groupFilters.createdFrom);
  const createdTo = clean(state.groupFilters.createdTo);
  const dateFrom = clean(state.groupFilters.dateFrom);
  const dateTo = clean(state.groupFilters.dateTo);
  const searchNeedle = clean(state.groupFilters.search).toLowerCase();
  return state.groups
    .filter((row) => !state.groupsShowActive || clean(row.checkOut) >= today)
    .filter((row) => {
      const created = clean(row.creationDate).slice(0, 10);
      if (createdFrom && (!created || created < createdFrom)) return false;
      if (createdTo && (!created || created > createdTo)) return false;
      return true;
    })
    .filter((row) => {
      const checkIn = clean(row.checkIn);
      const checkOut = clean(row.checkOut);
      if (dateFrom && checkOut && checkOut < dateFrom) return false;
      if (dateTo && checkIn && checkIn > dateTo) return false;
      return true;
    })
    .filter((row) => {
      if (!searchNeedle) return true;
      return clean(row.name).toLowerCase().includes(searchNeedle) || clean(row.reservationNumber).toLowerCase().includes(searchNeedle);
    })
    .sort(compareGroupRows);
}

function onGroupFilterInput() {
  state.groupFilters.createdFrom = clean(els.groupsFilterCreatedFrom?.value);
  state.groupFilters.createdTo = clean(els.groupsFilterCreatedTo?.value);
  state.groupFilters.dateFrom = clean(els.groupsFilterDateFrom?.value);
  state.groupFilters.dateTo = clean(els.groupsFilterDateTo?.value);
  state.groupFilters.search = clean(els.groupsFilterSearch?.value);
  renderGroups();
}

function compareGroupRows(a, b) {
  const dir = state.groupSort.dir === "desc" ? -1 : 1;
  const key = state.groupSort.key;
  const av = key === "created" ? clean(a.creationDate) : clean(a.checkIn);
  const bv = key === "created" ? clean(b.creationDate) : clean(b.checkIn);
  const primary = av.localeCompare(bv);
  if (primary) return primary * dir;
  return clean(a.name).localeCompare(clean(b.name));
}

function onGroupSortToggle(event) {
  const button = event.target.closest("button[data-group-sort]");
  if (!button) return;
  const key = clean(button.dataset.groupSort);
  if (!key) return;
  if (state.groupSort.key === key) {
    state.groupSort.dir = state.groupSort.dir === "asc" ? "desc" : "asc";
  } else {
    state.groupSort.key = key;
    state.groupSort.dir = key === "created" ? "desc" : "asc";
  }
  renderGroups();
}

function updateGroupSortIndicators() {
  const table = els.groupsRows?.closest("table");
  if (!table) return;
  table.querySelectorAll("button[data-group-sort]").forEach((button) => {
    const key = clean(button.dataset.groupSort);
    const active = key === state.groupSort.key;
    const indicator = button.querySelector(".sort-indicator");
    button.classList.toggle("active", active);
    if (indicator) indicator.textContent = active ? (state.groupSort.dir === "asc" ? "↑" : "↓") : "";
  });
}

function groupRoomTypeSummary(items = []) {
  return items
    .filter((item) => clean(item.roomType))
    .map((item) => `${normalizeGroupRoomCount(item.roomCount)}x ${clean(item.roomType)}`)
    .join(", ");
}

function groupRoomTypeSummaryHtml(items = []) {
  const lines = items
    .filter((item) => clean(item.roomType))
    .map((item) => `<span>${escape(`${normalizeGroupRoomCount(item.roomCount)}x ${clean(item.roomType)}`)}</span>`);
  return lines.length ? lines.join("") : "-";
}

function groupRoomSelectionSummary(items = []) {
  return items
    .flatMap((item) => (Array.isArray(item.rooms) ? item.rooms : []))
    .map(clean)
    .filter(Boolean)
    .join(", ");
}

function groupPdfRoomTypeLines(items = []) {
  return (Array.isArray(items) ? items : [])
    .filter((item) => clean(item.roomType))
    .map((item) => `${normalizeGroupRoomCount(item.roomCount)}x ${clean(item.roomType)}`);
}

function groupPdfRoomLines(items = []) {
  return (Array.isArray(items) ? items : [])
    .filter((item) => clean(item.roomType))
    .map((item) => {
      const rooms = (Array.isArray(item.rooms) ? item.rooms : []).map(clean).filter(Boolean);
      return rooms.length ? rooms.join(", ") : "-";
    });
}

function groupDepositText(totalValue) {
  const percentage = Number(state.groupSettings.depositPercentage || 0);
  return `Deposit ${percentage}%: ${formatMoney(Number(totalValue || 0) * (percentage / 100))}`;
}

function groupExportRows() {
  return getFilteredGroups().map((row) => ({
    created: formatDateOnly(row.creationDate),
    name: row.name,
    email: row.email,
    checkIn: formatGroupDateDisplay(row.checkIn),
    checkOut: formatGroupDateDisplay(row.checkOut),
    nights: dateDiffDays(row.checkIn, row.checkOut),
    guests: row.guests || 0,
    status: row.status,
    language: proposalLanguageLabel(row.language),
    total: formatMoney(row.totalValue),
    deposit: groupDepositText(row.totalValue),
    roomItems: Array.isArray(row.roomItems) ? row.roomItems : [],
    roomTypes: groupRoomTypeSummary(row.roomItems),
    rooms: groupRoomSelectionSummary(row.roomItems),
    optionDate: row.optionDate || "-",
    reservationNumber: row.reservationNumber || "-",
    observation: row.observation || "",
  }));
}

function groupFilterSummaryParts() {
  const parts = [state.groupsShowActive ? "Active only" : "All proposals"];
  if (clean(state.groupFilters.createdFrom) || clean(state.groupFilters.createdTo)) {
    parts.push(`Created: ${clean(state.groupFilters.createdFrom) || "-"} to ${clean(state.groupFilters.createdTo) || "-"}`);
  }
  if (clean(state.groupFilters.dateFrom) || clean(state.groupFilters.dateTo)) {
    parts.push(`Dates: ${clean(state.groupFilters.dateFrom) || "-"} to ${clean(state.groupFilters.dateTo) || "-"}`);
  }
  if (clean(state.groupFilters.search)) {
    parts.push(`Search: ${clean(state.groupFilters.search)}`);
  }
  return parts;
}

function exportGroupsToExcel() {
  const rows = groupExportRows();
  const headers = ["Created", "Name", "Email", "Check-in", "Check-out", "Nights", "Guests", "Status", "Language", "Total", "Deposit", "Room Types", "Rooms", "Option", "Reservation", "Observation"];
  const htmlRows = rows.map((row) => [
    row.created,
    row.name,
    row.email,
    row.checkIn,
    row.checkOut,
    row.nights,
    row.guests,
    row.status,
    row.language,
    row.total,
    row.deposit,
    row.roomTypes,
    row.rooms,
    row.optionDate,
    row.reservationNumber,
    row.observation,
  ]);
  const table = `<table border="1">
    <thead><tr>${headers.map((header) => `<th>${escape(header)}</th>`).join("")}</tr></thead>
    <tbody>${htmlRows.map((cells) => `<tr>${cells.map((cell) => `<td>${escape(cell)}</td>`).join("")}</tr>`).join("")}</tbody>
  </table>`;
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body>${table}</body></html>`;
  const date = formatDate(new Date());
  downloadBlob(`group_proposals_${date}.xls`, html, "application/vnd.ms-excel;charset=utf-8;");
  showToast(`Exported ${rows.length} group proposals to Excel.`, "success");
}

function exportGroupsToPdf() {
  const rows = groupExportRows();
  const date = formatGroupDateDisplay(formatDate(new Date()));
  const filterSummary = groupFilterSummaryParts().join(" · ");
  const tableRows = rows.map((row) => {
    const roomTypeLines = groupPdfRoomTypeLines(row.roomItems);
    const roomLines = groupPdfRoomLines(row.roomItems);
    return `<tr class="${row.status === "Accepted" ? "accepted" : row.status === "Refused" ? "refused" : ""}">
    <td>${escape(row.created)}</td>
    <td>${escape(row.name)}</td>
    <td>${escape(row.checkIn)}<br>${escape(row.checkOut)}<br><small>${escape(String(row.nights))} nights</small></td>
    <td>${escape(String(row.guests))}</td>
    <td><strong>${escape(row.total)}</strong><br><small>${escape(row.deposit)}</small></td>
    <td>${(roomTypeLines.length ? roomTypeLines : [row.roomTypes || "-"]).map((line) => escape(line)).join("<br>")}</td>
    <td>${(roomLines.length ? roomLines : [row.rooms || "-"]).map((line) => escape(line)).join("<br>")}</td>
    <td>${escape(row.optionDate)}</td>
    <td>${escape(row.reservationNumber)}</td>
  </tr>`;
  }).join("");
  const html = `<!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>Group Proposals</title>
        <style>
          @page { size: landscape; margin: 12mm; }
          body { font-family: Calibri, Arial, sans-serif; color: #1f1f1f; }
          .toolbar { display: flex; gap: 8px; align-items: center; margin: 0 0 16px; padding: 10px; background: #f6efe8; border: 1px solid #d8c8b8; border-radius: 10px; }
          .toolbar button { background: #0a5f57; color: white; border: 0; border-radius: 8px; padding: 8px 12px; font-weight: 700; cursor: pointer; }
          .toolbar span { color: #5f554c; font-size: 13px; }
          h1 { margin: 0 0 4px; font-size: 22px; }
          p { margin: 0 0 14px; color: #666; }
          .filters { margin-top: -8px; margin-bottom: 14px; color: #5f554c; font-size: 12px; }
          body > p:not(.summary):not(.filters) { display: none; }
          table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11px; }
          th { background: #0a5f57; color: white; border: 1px solid #0a5f57; padding: 6px; text-align: left; }
          td { border: 1px solid #cfc7bd; padding: 6px; vertical-align: top; word-wrap: break-word; }
          tr.accepted td { background: rgba(46, 159, 66, 0.22); }
          tr.refused td { background: rgba(212, 76, 76, 0.22); }
          small { color: #555; }
          @media print { .toolbar { display: none; } }
        </style>
      </head>
      <body>
        <div class="toolbar">
          <button type="button" onclick="window.print()">Print / Save PDF</button>
          <span>If the print dialog does not open automatically, press this button and choose "Save as PDF".</span>
        </div>
        <h1>Group Proposals</h1>
        <p class="summary">Exported ${escape(date)} &middot; ${escape(String(rows.length))} proposal${rows.length === 1 ? "" : "s"}</p>
        <p class="filters"><strong>Filters:</strong> ${escape(filterSummary || "None")}</p>
        <p>Exported ${escape(date)} · ${escape(String(rows.length))} proposal${rows.length === 1 ? "" : "s"} · ${state.groupsShowActive ? "Active only" : "All proposals"}</p>
        <table>
          <colgroup>
            <col style="width: 8%">
            <col style="width: 14%">
            <col style="width: 11%">
            <col style="width: 5%">
            <col style="width: 9%">
            <col style="width: 28%">
            <col style="width: 8%">
            <col style="width: 7%">
            <col style="width: 10%">
          </colgroup>
          <thead>
            <tr><th>Created</th><th>Name</th><th>Dates</th><th>Guests</th><th>Total</th><th>Room Types</th><th>Rooms</th><th>Option</th><th>Reservation</th></tr>
          </thead>
          <tbody>${tableRows || '<tr><td colspan="9">No group proposals found.</td></tr>'}</tbody>
        </table>
        <script>
          window.addEventListener("load", () => {
            window.focus();
            setTimeout(() => window.print(), 700);
          });
        </script>
      </body>
    </html>`;
  const win = window.open("", "_blank");
  if (!win) {
    showToast("Could not open PDF print window. Please allow pop-ups for this site.", "error");
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
}

function groupMetadataObservation(value) {
  return clean(groupMetadata(value).observation);
}

function groupMetadata(value) {
  const items = Array.isArray(value) ? value : [];
  const metadata = items.find((item) => item && item.type === "metadata") || {};
  return {
    observation: clean(metadata.observation),
    language: normalizeProposalLanguage(metadata.language),
    audit: normalizeGroupAudit(metadata.audit),
  };
}

function normalizeGroupAudit(value) {
  return (Array.isArray(value) ? value : [])
    .map((item) => ({
      at: clean(item?.at),
      action: clean(item?.action),
      summary: clean(item?.summary),
    }))
    .filter((item) => item.at && item.action)
    .slice(-20);
}

function buildGroupMetadata(draft, audit = draft.audit) {
  const observation = clean(draft.observation);
  const metadata = {
    type: "metadata",
    observation,
    language: normalizeProposalLanguage(draft.language || state.groupProposalLanguage),
    audit: normalizeGroupAudit(audit),
  };
  return [metadata];
}

function renderGroupDraft() {
  const draft = state.groupDraft;
  state.groupProposalLanguage = normalizeProposalLanguage(draft.language || state.groupProposalLanguage);
  renderGroupEditorTab();
  els.groupReservationNumber.value = draft.reservationNumber;
  els.groupName.value = draft.name;
  els.groupEmail.value = draft.email;
  renderGroupEmailProposalHint();
  els.groupCheckIn.value = draft.checkIn || "";
  els.groupCheckOut.value = draft.checkOut || "";
  syncGroupDateConstraints();
  els.groupGuests.value = draft.guests;
  els.groupOptionDate.value = draft.optionDate;
  els.groupStatusField.value = draft.status;
  renderGroupStatusColor();
  els.groupObservation.value = draft.observation;
  renderGroupRoomItems();
  renderGroupTotals();
  renderGroupAuditHistory();
  els.groupDelete.hidden = !draft.id || !isAdministratorProfile();
}

function nextGroupCheckOutDate(checkIn) {
  const raw = clean(checkIn);
  if (!raw) return "";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return "";
  date.setDate(date.getDate() + 1);
  return formatDate(date);
}

function syncGroupDateConstraints() {
  const minCheckOut = nextGroupCheckOutDate(state.groupDraft.checkIn);
  if (els.groupCheckOut) els.groupCheckOut.min = minCheckOut || "";
}

function prepareGroupCheckOutPicker() {
  syncGroupDateConstraints();
  const minCheckOut = clean(els.groupCheckOut?.min || nextGroupCheckOutDate(state.groupDraft.checkIn));
  if (!minCheckOut) return;
  const current = parseGroupDateInput(els.groupCheckOut?.value || state.groupDraft.checkOut);
  if (!current || current < minCheckOut) {
    els.groupCheckOut.value = minCheckOut;
    state.groupDraft.checkOut = minCheckOut;
    renderGroupTotals();
  }
}

function renderGroupTotals() {
  const draft = state.groupDraft;
  syncGroupRoomItemGuests(draft);
  const accommodationTotal = calculateGroupAccommodationTotal(draft);
  const cityTaxTotal = calculateGroupCityTaxTotal(draft);
  const total = accommodationTotal + cityTaxTotal;
  draft.totalValue = total;
  els.groupAccommodationTotal.textContent = formatMoney(accommodationTotal);
  els.groupCityTaxTotal.textContent = formatMoney(cityTaxTotal);
  els.groupTotalValue.textContent = formatMoney(total);
  const deposit = total * (Number(state.groupSettings.depositPercentage || 0) / 100);
  els.groupDepositPreview.textContent = `Deposit (${state.groupSettings.depositPercentage || 0}%): ${formatMoney(deposit)}`;
  els.groupNightsLabel.textContent = groupNightsLabelText(draft.checkIn, draft.checkOut);
  els.groupLastPaymentLimit.value = groupLastPaymentLimitText(draft.checkIn);
  const remaining = groupRemainingGuests(draft);
  els.groupGuestCounter.textContent = remaining < 0 ? `Guests over: ${Math.abs(remaining)}` : `Guests remaining: ${remaining}`;
  els.groupGuestCounter.classList.toggle("overbooked", remaining < 0);
  (draft.roomItems || []).forEach((item, index) => {
    const guestsEl = els.groupRoomItemsBody.querySelector(`[data-group-room-guests-display="${index}"]`);
    const lineTotalEl = els.groupRoomItemsBody.querySelector(`[data-group-room-line-total="${index}"]`);
    if (guestsEl) guestsEl.textContent = String(item.guests || 0);
    if (lineTotalEl) lineTotalEl.textContent = formatMoney(calculateGroupRoomItemTotal(draft, item));
  });
  renderGroupProposalEmail();
}

function renderGroupAuditHistory() {
  if (!els.groupAuditHistory) return;
  const audit = normalizeGroupAudit(state.groupDraft.audit);
  if (!audit.length) {
    els.groupAuditHistory.classList.add("empty");
    els.groupAuditHistory.innerHTML = "No saved changes yet.";
    return;
  }
  els.groupAuditHistory.classList.remove("empty");
  els.groupAuditHistory.innerHTML = audit
    .slice()
    .reverse()
    .map((item) => `<article><strong>${escape(item.action)}</strong><span>${escape(formatDateTimeShort(item.at))}</span><p>${escape(item.summary || "-")}</p></article>`)
    .join("");
}

function renderGroupEditorTab() {
  const isDetails = state.groupEditorTab === "details";
  const isProposal = state.groupEditorTab === "email";
  const isConfirmation = state.groupEditorTab === "confirmation";
  const isFinalConfirmation = state.groupEditorTab === "final-confirmation";
  els.groupTabDetails.classList.toggle("active-tab", isDetails);
  els.groupTabDetails.classList.toggle("ghost", !isDetails);
  els.groupTabEmail.classList.toggle("active-tab", isProposal);
  els.groupTabEmail.classList.toggle("ghost", !isProposal);
  els.groupTabConfirmation.classList.toggle("active-tab", isConfirmation);
  els.groupTabConfirmation.classList.toggle("ghost", !isConfirmation);
  els.groupTabFinalConfirmation.classList.toggle("active-tab", isFinalConfirmation);
  els.groupTabFinalConfirmation.classList.toggle("ghost", !isFinalConfirmation);
  els.groupDetailsPanel.hidden = !isDetails;
  els.groupEmailPanel.hidden = isDetails;
}

function setGroupEditorTab(tab) {
  state.groupEditorTab = ["email", "confirmation", "final-confirmation"].includes(tab) ? tab : "details";
  renderGroupEditorTab();
  renderGroupProposalEmail();
}

async function copyGroupEmailText() {
  const confirmationKind = groupConfirmationKind();
  const text = confirmationKind ? groupConfirmationEmailText(state.groupDraft, confirmationKind) : groupProposalEmailText(state.groupDraft);
  const html = confirmationKind ? groupConfirmationEmailHtml(state.groupDraft, confirmationKind) : groupProposalEmailHtml(state.groupDraft);
  const label = confirmationKind ? "confirmation" : "proposal";
  try {
    if (window.ClipboardItem) {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" }),
        }),
      ]);
      setGroupsStatus(`Formatted ${label} email copied.`);
      showToast(`Formatted ${label} email copied.`, "success");
    } else {
      await navigator.clipboard.writeText(text);
      setGroupsStatus(`Plain ${label} email text copied.`);
      showToast(`Plain ${label} email text copied.`, "success");
    }
  } catch (e) {
    setGroupsStatus("Could not copy automatically. Select the preview and copy it manually.");
  }
}

function renderGroupProposalEmail() {
  if (!els.groupEmailPreview) return;
  const confirmationKind = groupConfirmationKind();
  els.groupEmailTitle.textContent = confirmationKind === "final" ? "Final confirmation Text" : confirmationKind ? "1st confirmation Text" : "Proposal Text";
  els.groupEmailDescription.textContent = confirmationKind === "final"
    ? "Generated from the proposal and the configurable final confirmation template in Groups settings."
    : confirmationKind
    ? "Generated from the proposal and the configurable 1st confirmation template in Groups settings."
    : "Generated from the proposal and the configurable template in Groups settings.";
  els.groupProposalLanguage.closest("label").hidden = false;
  els.groupCopyEmail.textContent = confirmationKind ? "Copy Confirmation Text" : "Copy Email Text";
  els.groupProposalLanguage.value = normalizeProposalLanguage(state.groupProposalLanguage);
  els.groupEmailPreview.innerHTML = confirmationKind
    ? groupConfirmationEmailHtml(state.groupDraft, confirmationKind)
    : groupProposalEmailHtml(state.groupDraft);
}

function groupProposalEmailText(draft) {
  const template = groupProposalTemplate();
  const replacements = groupProposalEmailReplacements(draft);
  return Object.entries(replacements).reduce(
    (text, [key, value]) => text.replaceAll(`{{${key}}}`, value),
    template
  );
}

function groupProposalEmailHtml(draft) {
  return groupProposalEmailHtmlFromTemplate(draft, groupProposalTemplate());
}

function groupProposalEmailHtmlFromTemplate(draft, template) {
  const replacements = groupProposalEmailReplacements(draft);
  const textWithValues = Object.entries(replacements)
    .filter(([key]) => key !== "room_table")
    .reduce((text, [key, value]) => text.replaceAll(`{{${key}}}`, value), template);
  const chunks = stripProposalTotalLines(textWithValues).split("{{room_table}}").map(groupProposalTextChunkHtml);
  const tableAndTotals = `${groupProposalRoomTableHtml(draft)}${groupProposalTotalsHtml(draft)}`;
  return `<div class="proposal-email-document" style="font-family: Calibri, Arial, Helvetica, sans-serif; color: #000000; font-size: 11pt; line-height: 1.15; max-width: 680px;">${chunks.join(tableAndTotals)}</div>`;
}

function groupConfirmationKind() {
  if (state.groupEditorTab === "confirmation") return "first";
  if (state.groupEditorTab === "final-confirmation") return "final";
  return "";
}

function groupConfirmationEmailText(draft, kind = "first") {
  const template = groupConfirmationTemplate(kind);
  const replacements = groupProposalEmailReplacements(draft, kind);
  return Object.entries(replacements).reduce(
    (text, [key, value]) => text.replaceAll(`{{${key}}}`, value),
    template
  );
}

function groupConfirmationEmailHtml(draft, kind = "first") {
  return groupConfirmationEmailHtmlFromTemplate(draft, groupConfirmationTemplate(kind), kind);
}

function groupConfirmationEmailHtmlFromTemplate(draft, template, kind = "first") {
  const replacements = groupProposalEmailReplacements(draft, kind);
  const textWithValues = Object.entries(replacements)
    .filter(([key]) => key !== "confirmation_table")
    .reduce((text, [key, value]) => text.replaceAll(`{{${key}}}`, value), template);
  const chunks = textWithValues.split("{{confirmation_table}}").map(groupProposalTextChunkHtml);
  return `<div class="proposal-email-document" style="font-family: Calibri, Arial, Helvetica, sans-serif; color: #000000; font-size: 11pt; line-height: 1.15; max-width: 680px;">${chunks.join(groupConfirmationTableHtml(draft, kind))}</div>`;
}

function groupProposalTemplate() {
  const language = normalizeProposalLanguage(state.groupProposalLanguage);
  if (language === "pt" || language === "es") return GROUP_PROPOSAL_TEMPLATES[language];
  return clean(state.groupSettings.emailTemplate) || DEFAULT_GROUP_SETTINGS.emailTemplate;
}

function groupConfirmationTemplate(kind = "first") {
  const language = normalizeProposalLanguage(state.groupProposalLanguage);
  if ((language === "pt" || language === "es") && GROUP_CONFIRMATION_TEMPLATES[kind]?.[language]) {
    return GROUP_CONFIRMATION_TEMPLATES[kind][language];
  }
  if (kind === "final") {
    return clean(state.groupSettings.finalConfirmationTemplate) || DEFAULT_GROUP_SETTINGS.finalConfirmationTemplate;
  }
  return clean(state.groupSettings.confirmationTemplate) || DEFAULT_GROUP_SETTINGS.confirmationTemplate;
}

function normalizeProposalLanguage(value) {
  const raw = clean(value).toLowerCase();
  return raw === "pt" || raw === "es" ? raw : "en";
}

function proposalLanguageLabel(value) {
  const language = normalizeProposalLanguage(value);
  if (language === "pt") return "Portuguese";
  if (language === "es") return "Spanish";
  return "English";
}

function stripProposalTotalLines(text) {
  return String(text || "")
    .split("\n")
    .filter((line) => !/^\s*(?:Accommod\.|Accommodation)\s+Total\s*=/.test(line))
    .filter((line) => !/^\s*Total\s+(?:Alojamento|Alojamiento)\s*=/.test(line))
    .filter((line) => !/^\s*City Tax\b.*=/.test(line))
    .filter((line) => !/^\s*(?:Taxa|Tasa)\s+Municipal\b.*=/.test(line))
    .filter((line) => !/^\s*Total\s*=/.test(line))
    .join("\n");
}

function groupProposalTextChunkHtml(text) {
  return clean(text)
    .split(/\n{2,}/)
    .filter(Boolean)
    .map(proposalParagraphHtml)
    .join("");
}

function proposalParagraphHtml(paragraph) {
  const raw = clean(paragraph);
  if (/^(Payment can be made by:|O pagamento pode ser feito por:|El pago se puede realizar mediante:)/i.test(raw) && raw.includes("\n")) {
    const lines = raw.split("\n").map(clean).filter(Boolean);
    const intro = lines.shift();
    const items = lines.map((line) => `<li style="margin: 0 0 2px;">${proposalInlineFormatHtml(line)}</li>`).join("");
    return `<p style="margin: 0 0 4pt 47px; padding: 0 0 3pt 0; text-align: justify; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; line-height: 1.15; color: #000000; mso-line-height-rule: exactly;">${proposalInlineFormatHtml(intro)}</p><ul style="margin: 0 0 10pt 70px; padding: 0; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; line-height: 1.15; color: #000000; mso-line-height-rule: exactly;">${items}</ul><div style="height: 6pt; line-height: 6pt; font-size: 6pt;">&nbsp;</div>`;
  }
  const isTotalLine = /^(Accommod\. Total|Accommodation Total|Total Alojamento|Total Alojamiento|City Tax|Taxa Municipal|Tasa Municipal|Total\s*=)/i.test(raw);
  const isCityTaxNote = /^(Please note that there is a city tax|Por favor note que existe uma taxa municipal|Tenga en cuenta que existe una tasa municipal)/i.test(raw);
  const style = [
    "margin: 0 0 10pt 47px",
    "padding: 0 0 3pt 0",
    "text-align: justify",
    "font-family: Calibri, Arial, Helvetica, sans-serif",
    "font-size: 11pt",
    "line-height: 1.15",
    "mso-line-height-rule: exactly",
    "color: #000000",
    isTotalLine || isCityTaxNote ? "font-weight: bold" : "",
    isCityTaxNote ? "font-style: italic" : "",
  ].filter(Boolean).join("; ");
  return `<p style="${style};">${proposalInlineFormatHtml(raw)}</p><div style="height: 4pt; line-height: 4pt; font-size: 4pt;">&nbsp;</div>`;
}

function proposalInlineFormatHtml(text) {
  let html = escape(text).replace(/\n/g, "<br>");
  [
    "free breakfast",
    "8:00 AM to 11:00 AM",
    "(three types of cereals, three types of bread, muffins, mini croissants, jam, honey, butter, peanut butter, chocolate cream, fruit, coffee, tea, cocoa, milk, juice and our homemade pancakes!)",
    "pequeno-almoço gratuito",
    "08:00 às 11:00",
    "três tipos de cereais, três tipos de pão, muffins, mini croissants, compota, mel, manteiga, manteiga de amendoim, creme de chocolate, fruta, café, chá, cacau, leite, sumo e as nossas panquecas caseiras!",
    "desayuno gratuito",
    "08:00 a 11:00",
    "tres tipos de cereales, tres tipos de pan, muffins, mini croissants, mermelada, miel, mantequilla, mantequilla de cacahuete, crema de chocolate, fruta, café, té, cacao, leche, zumo y nuestras tortitas caseras.",
  ].forEach((phrase) => {
    html = html.replaceAll(escape(phrase), `<strong>${escape(phrase)}</strong>`);
  });
  return html;
}

function groupProposalEmailReplacements(draft, confirmationKind = "") {
  const accommodationTotal = calculateGroupAccommodationTotal(draft);
  const cityTaxTotal = calculateGroupCityTaxTotal(draft);
  const total = accommodationTotal + cityTaxTotal;
  const depositPercentage = Number(state.groupSettings.depositPercentage || 0);
  const depositValue = total * (depositPercentage / 100);
  const cityTaxNights = groupCityTaxableNights(draft);
  return {
    name: clean(draft.name) || "[name]",
    arrival: formatGroupDateDisplay(draft.checkIn),
    departure: formatGroupDateDisplay(draft.checkOut),
    nights: String(dateDiffDays(draft.checkIn, draft.checkOut)),
    guests: String(normalizeGroupGuests(draft.guests, 0)),
    city_tax_guests: String(groupCityTaxableGuests(draft)),
    room_table: groupProposalRoomTableText(draft),
    accommodation_total: formatMoney(accommodationTotal),
    city_tax_nights: String(cityTaxNights),
    city_tax_total: formatMoney(cityTaxTotal),
    total: formatMoney(total),
    deposit_percentage: String(depositPercentage),
    deposit_value: formatMoney(depositValue),
    balance_due: formatMoney(Math.max(0, total - depositValue)),
    last_payment_days: String(normalizeLastPaymentDays(state.groupSettings.lastPaymentDaysBeforeArrival)),
    last_payment_date: groupLastPaymentLimitText(draft.checkIn),
    option_date: draft.optionDate ? formatGroupDateDisplay(draft.optionDate) : "-",
    reservation_number: clean(draft.reservationNumber) || "-",
    rooms_booked: groupBookedRoomsText(draft),
    confirmation_table: groupConfirmationTableText(draft, confirmationKind || groupConfirmationKind() || "first"),
  };
}

function groupBookedRoomsList(draft) {
  const items = (draft.roomItems || []).filter((item) => clean(item.roomType));
  return items.map((item) => {
    const roomCount = normalizeGroupRoomCount(item.roomCount);
    const subgroup = clean(item.subgroup);
    const tags = [subgroup, item.under13 ? groupUnder13Label() : ""].filter(Boolean).join(", ");
    return `${roomCount}x ${groupProposalRoomTypeName(item.roomType)}${tags ? ` (${tags})` : ""}`;
  });
}

function groupBookedRoomsText(draft) {
  const rooms = groupBookedRoomsList(draft);
  return rooms.length ? rooms.join("\n") : "-";
}

function groupConfirmationTableRows(draft, kind = "first") {
  const labels = groupConfirmationLabels();
  const accommodationTotal = calculateGroupAccommodationTotal(draft);
  const cityTaxTotal = calculateGroupCityTaxTotal(draft);
  const total = accommodationTotal + cityTaxTotal;
  const depositPercentage = Number(state.groupSettings.depositPercentage || 0);
  const depositValue = total * (depositPercentage / 100);
  const rows = [
    [labels.bookingRef, clean(draft.reservationNumber) || "-"],
    [labels.roomsBooked, groupBookedRoomsText(draft)],
    [labels.guests, String(normalizeGroupGuests(draft.guests, 0))],
    [labels.arrival, formatGroupDateDisplay(draft.checkIn)],
    [labels.nights, String(dateDiffDays(draft.checkIn, draft.checkOut))],
    [labels.departure, formatGroupDateDisplay(draft.checkOut)],
    [labels.accommodationTotal, formatMoney(accommodationTotal)],
    [labels.cityTaxTotal, formatMoney(cityTaxTotal)],
    [labels.totalWithTax, formatMoney(total)],
  ];
  if (kind === "final") {
    rows.push([labels.totalPaid, formatMoney(total)]);
    return rows;
  }
  rows.push([`${labels.depositPaid} ${depositPercentage}% (${labels.nonRefundable})`, formatMoney(depositValue)]);
  rows.push([`${labels.balanceDueUntil} ${groupLastPaymentLimitText(draft.checkIn)} (${normalizeLastPaymentDays(state.groupSettings.lastPaymentDaysBeforeArrival)} ${labels.daysBeforeArrival})`, formatMoney(Math.max(0, total - depositValue))]);
  return rows;
}

function groupConfirmationTableText(draft, kind = "first") {
  return groupConfirmationTableRows(draft, kind)
    .map(([label, value]) => `${label}: ${value}`)
    .join("\n");
}

function groupConfirmationLabels() {
  const language = normalizeProposalLanguage(state.groupProposalLanguage);
  if (language === "pt") {
    return {
      bookingRef: "Ref. da reserva",
      roomsBooked: "Tipo de quartos reservados",
      guests: "Nº de hóspedes",
      arrival: "Chegada",
      nights: "Nº de noites",
      departure: "Partida",
      accommodationTotal: "Total alojamento",
      cityTaxTotal: "Total taxa municipal",
      totalWithTax: "Total da reserva com taxa municipal",
      depositPaid: "Depósito pago",
      nonRefundable: "não reembolsável",
      balanceDueUntil: "Total a pagar até",
      daysBeforeArrival: "dias antes da chegada",
      totalPaid: "Total pago",
    };
  }
  if (language === "es") {
    return {
      bookingRef: "Ref. de reserva",
      roomsBooked: "Tipo de habitaciones reservadas",
      guests: "Nº de huéspedes",
      arrival: "Llegada",
      nights: "Nº de noches",
      departure: "Salida",
      accommodationTotal: "Total alojamiento",
      cityTaxTotal: "Total tasa municipal",
      totalWithTax: "Total de la reserva con tasa municipal",
      depositPaid: "Depósito pagado",
      nonRefundable: "no reembolsable",
      balanceDueUntil: "Total a pagar hasta",
      daysBeforeArrival: "días antes de la llegada",
      totalPaid: "Total pagado",
    };
  }
  return {
    bookingRef: "Booking Ref",
    roomsBooked: "Type of rooms booked",
    guests: "Nr of Guests",
    arrival: "Arrival",
    nights: "Nr of Nights",
    departure: "Departure",
    accommodationTotal: "Total Accommodation",
    cityTaxTotal: "Total City Tax",
    totalWithTax: "Total of the reservation with City TAX",
    depositPaid: "Deposit Paid",
    nonRefundable: "non refundable",
    balanceDueUntil: "Total to be paid until",
    daysBeforeArrival: "days before arrival",
    totalPaid: "Total paid",
  };
}

function groupConfirmationTableHtml(draft, kind = "first") {
  const labelStyle = "border: 1pt solid #4F81BD; background: #D9EAF7; padding: 5px 8px; width: 210px; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; font-weight: bold; color: #000000; vertical-align: top;";
  const valueStyle = "border: 1pt solid #4F81BD; padding: 5px 8px; width: 250px; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; color: #000000; vertical-align: top;";
  const rows = groupConfirmationTableRows(draft, kind)
    .map(([label, value]) => `<tr><td style="${labelStyle}">${escape(label)}</td><td style="${valueStyle}">${escape(value).replace(/\n/g, "<br>")}</td></tr>`)
    .join("");
  return `<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin: 12px 0 18px 47px; width: 460px; table-layout: fixed; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; border: 1pt solid #4F81BD;"><tbody>${rows}</tbody></table>`;
}

function groupProposalRoomTableText(draft) {
  syncGroupRoomItemGuests(draft);
  const labels = groupProposalLabels();
  const nights = dateDiffDays(draft.checkIn, draft.checkOut);
  const lines = [
    `${labels.roomType} | ${labels.price} | ${labels.guests} | ${labels.nights} | ${labels.total}`,
    "------------------------------------------------------",
  ];
  const items = (draft.roomItems || []).filter((item) => clean(item.roomType));
  if (!items.length) {
    lines.push("[Add room types and prices in the proposal]");
    return lines.join("\n");
  }
  items.forEach((item) => {
    const label = groupProposalRoomLabel(item);
    const priceLabel = `${formatMoney(normalizeNumber(item.price))} ${item.priceMode === "room" ? labels.perRoom : labels.perGuest}`;
    lines.push(`${label} | ${priceLabel} | ${normalizeGroupGuests(item.guests, 0)} | ${nights} | ${formatMoney(calculateGroupRoomItemTotal(draft, item))}`);
  });
  return lines.join("\n");
}

function groupProposalLabels() {
  const language = normalizeProposalLanguage(state.groupProposalLanguage);
  if (language === "pt") {
    return {
      roomType: "Tipo de quarto",
      price: "Preço",
      guests: "Nº Hóspedes",
      nights: "Nº Noites",
      total: "Total",
      perGuest: "por hóspede",
      perRoom: "por quarto",
      accommodationTotal: "Total Alojamento",
      cityTax: "Taxa Municipal",
      guestWord: "hóspedes",
      nightWord: "noites",
      emptyTable: "Adicione tipos de quarto e preços na proposta.",
    };
  }
  if (language === "es") {
    return {
      roomType: "Tipo de habitación",
      price: "Precio",
      guests: "Nº Huéspedes",
      nights: "Nº Noches",
      total: "Total",
      perGuest: "por huésped",
      perRoom: "por habitación",
      accommodationTotal: "Total Alojamiento",
      cityTax: "Tasa Municipal",
      guestWord: "huéspedes",
      nightWord: "noches",
      emptyTable: "Añada tipos de habitación y precios en la propuesta.",
    };
  }
  return {
    roomType: "Bedroom types",
    price: "Price",
    guests: "Nº Guests",
    nights: "Nº Nights",
    total: "Total",
    perGuest: "per guest",
    perRoom: "per room",
    accommodationTotal: "Accommodation Total",
    cityTax: "City Tax",
    guestWord: "guests",
    nightWord: "nights",
    emptyTable: "Add room types and prices in the proposal.",
  };
}

function groupProposalRoomLabel(item) {
  const roomCount = normalizeGroupRoomCount(item.roomCount);
  const subgroup = clean(item.subgroup);
  const tags = [subgroup, item.under13 ? groupUnder13Label() : ""].filter(Boolean).join(", ");
  const suffix = tags ? ` (${tags})` : "";
  return `${roomCount}x ${groupProposalRoomTypeName(item.roomType)}${suffix}`;
}

function groupProposalRoomLabelHtml(item) {
  const roomCount = normalizeGroupRoomCount(item.roomCount);
  const subgroup = clean(item.subgroup);
  const tags = [subgroup, item.under13 ? groupUnder13Label() : ""].filter(Boolean).join(", ");
  const base = `${roomCount}x ${groupProposalRoomTypeName(item.roomType)}`;
  return `${escape(base)}${tags ? ` <span style="font-weight: normal;">(${escape(tags)})</span>` : ""}`;
}

function groupProposalRoomTypeName(roomType) {
  const name = clean(roomType);
  const language = normalizeProposalLanguage(state.groupProposalLanguage);
  return GROUP_ROOM_TYPE_TRANSLATIONS[language]?.[name] || name;
}

function groupUnder13Label() {
  const language = normalizeProposalLanguage(state.groupProposalLanguage);
  if (language === "pt") return "menores de 13";
  if (language === "es") return "menores de 13";
  return "under 13";
}

function groupProposalRoomTableHtml(draft) {
  syncGroupRoomItemGuests(draft);
  const labels = groupProposalLabels();
  const nights = dateDiffDays(draft.checkIn, draft.checkOut);
  const items = (draft.roomItems || []).filter((item) => clean(item.roomType));
  if (!items.length) return `<p class="proposal-empty-table" style="margin: 0 0 12px 47px; color: #7a5b25; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt;">${escape(labels.emptyTable)}</p>`;
  const headerCellStyle = "border: 1pt solid #4F81BD; background: #4F81BD; padding: 0 5px; height: 33pt; text-align: center; vertical-align: middle; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; font-weight: bold; color: #FFFFFF;";
  const bodyCellStyle = "border: 1pt solid #4F81BD; padding: 0 5px; height: 31pt; text-align: center; vertical-align: middle; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; color: #000000;";
  const firstBodyCellStyle = `${bodyCellStyle} font-weight: bold;`;
  return `<table class="proposal-email-table" cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin: 12px 0 18px 47px; width: 558px; table-layout: fixed; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; border: 1pt solid #4F81BD;">
    <colgroup>
      <col style="width: 189px;">
      <col style="width: 129px;">
      <col style="width: 70px;">
      <col style="width: 76px;">
      <col style="width: 94px;">
    </colgroup>
    <thead>
      <tr>
        <th style="${headerCellStyle}">${escape(labels.roomType)}</th>
        <th style="${headerCellStyle}">${escape(labels.price)}</th>
        <th style="${headerCellStyle}">${escape(labels.guests)}</th>
        <th style="${headerCellStyle}">${escape(labels.nights)}</th>
        <th style="${headerCellStyle}">${escape(labels.total)}</th>
      </tr>
    </thead>
    <tbody>
      ${items.map((item) => {
        const priceLabel = `${formatMoney(normalizeNumber(item.price))} ${item.priceMode === "room" ? labels.perRoom : labels.perGuest}`;
        return `<tr>
          <td style="${firstBodyCellStyle}">${groupProposalRoomLabelHtml(item)}</td>
          <td style="${bodyCellStyle}">${escape(priceLabel)}</td>
          <td style="${bodyCellStyle}">${escape(String(normalizeGroupGuests(item.guests, 0)))}</td>
          <td style="${bodyCellStyle}">${escape(String(nights))}</td>
          <td style="${bodyCellStyle}">${escape(formatMoney(calculateGroupRoomItemTotal(draft, item)))}</td>
        </tr>`;
      }).join("")}
    </tbody>
  </table>`;
}

function groupProposalTotalsHtml(draft) {
  const labels = groupProposalLabels();
  const accommodationTotal = calculateGroupAccommodationTotal(draft);
  const cityTaxTotal = calculateGroupCityTaxTotal(draft);
  const total = accommodationTotal + cityTaxTotal;
  const cityTaxNights = groupCityTaxableNights(draft);
  const lineStyle = "margin: 0 0 8pt 47px; padding: 0 0 3pt 0; text-align: right; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; line-height: 1.15; mso-line-height-rule: exactly; font-weight: bold; color: #000000; width: 558px;";
  return `<div style="height: 4pt; line-height: 4pt; font-size: 4pt;">&nbsp;</div>
    <p style="${lineStyle}">${escape(labels.accommodationTotal)} = ${escape(formatMoney(accommodationTotal))}</p>
    <p style="${lineStyle}">${escape(labels.cityTax)} ${escape(String(groupCityTaxableGuests(draft)))} ${escape(labels.guestWord)} x ${escape(String(cityTaxNights))} ${escape(labels.nightWord)} x 4&euro; = ${escape(formatMoney(cityTaxTotal))}</p>
    <p style="${lineStyle}">${escape(labels.total)} = ${escape(formatMoney(total))}</p>
    <div style="height: 8pt; line-height: 8pt; font-size: 8pt;">&nbsp;</div>`;
}

function renderGroupRoomItems() {
  els.groupRoomItemsBody.innerHTML = "";
  if (!state.groupDraft.roomItems.length) {
    els.groupRoomItemsBody.innerHTML = '<tr><td colspan="10" class="empty">Add at least one room type.</td></tr>';
    return;
  }
  syncGroupRoomItemGuests(state.groupDraft);
  state.groupDraft.roomItems.forEach((item, index) => {
    const availableRooms = roomsForGroupRoomType(item.roomType);
    const roomsUsedElsewhere = selectedGroupRoomsExcept(index);
    const lineTotal = calculateGroupRoomItemTotal(state.groupDraft, item);
    const roomTypeOptions = ['<option value="">Select room type</option>']
      .concat(state.groupSettings.roomTypes.map((roomType) => `<option value="${escape(roomType.name)}">${escape(roomType.name)}</option>`))
      .join("");
    const roomOptions = availableRooms.map((room) => {
      const selected = item.rooms?.includes(room);
      const disabled = roomsUsedElsewhere.has(room) && !selected;
      return `<option value="${escape(room)}" ${selected ? "selected" : ""} ${disabled ? "disabled" : ""}>${escape(room)}</option>`;
    }).join("");
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input data-group-room-subgroup="${index}" value="${escape(clean(item.subgroup))}" placeholder="boys" maxlength="24" /></td>
      <td class="wide-col"><select data-group-room-type="${index}">${roomTypeOptions}</select></td>
      <td><input type="number" min="1" max="20" step="1" data-group-room-count="${index}" value="${escape(item.roomCount ?? 1)}" /></td>
      <td><strong data-group-room-guests-display="${index}">${escape(String(item.guests || 0))}</strong></td>
      <td><input type="checkbox" data-group-room-under13="${index}" ${item.under13 ? "checked" : ""} title="Exclude this line from city tax" /></td>
      <td><input type="number" min="0" step="0.01" data-group-room-price="${index}" value="${escape(item.price ?? "")}" /></td>
      <td><select data-group-room-mode="${index}">
        <option value="guest">Per guest</option>
        <option value="room">Per room</option>
      </select></td>
      <td><strong data-group-room-line-total="${index}">${escape(formatMoney(lineTotal))}</strong></td>
      <td><select multiple data-group-room-rooms="${index}">${roomOptions}</select></td>
      <td><button type="button" class="danger" data-remove-group-room="${index}">Remove</button></td>`;
    tr.querySelector(`[data-group-room-type="${index}"]`).value = clean(item.roomType);
    tr.querySelector(`[data-group-room-mode="${index}"]`).value = clean(item.priceMode) || "guest";
    els.groupRoomItemsBody.appendChild(tr);
  });
}

function onGroupDraftInput(event) {
  state.groupDraft.reservationNumber = clean(els.groupReservationNumber.value);
  state.groupDraft.name = clean(els.groupName.value);
  state.groupDraft.email = clean(els.groupEmail.value);
  state.groupDraft.checkIn = parseGroupDateInput(els.groupCheckIn.value);
  state.groupDraft.checkOut = parseGroupDateInput(els.groupCheckOut.value);
  syncGroupDateConstraints();
  state.groupDraft.guests = normalizeGroupGuests(els.groupGuests.value);
  state.groupDraft.optionDate = clean(els.groupOptionDate.value);
  state.groupDraft.status = normalizeGroupStatus(els.groupStatusField.value);
  state.groupDraft.observation = clean(els.groupObservation.value);
  renderGroupStatusColor();
  renderGroupTotals();
}

function addGroupRoomItem() {
  state.groupDraft.roomItems.push({ subgroup: "", roomType: "", roomCount: 1, guests: 0, under13: false, price: 0, priceMode: "guest", rooms: [] });
  renderGroupDraft();
}

function onGroupRoomItemInput(event) {
  const idx = Number(event.target.dataset.groupRoomSubgroup ?? event.target.dataset.groupRoomType ?? event.target.dataset.groupRoomCount ?? event.target.dataset.groupRoomUnder13 ?? event.target.dataset.groupRoomPrice ?? event.target.dataset.groupRoomMode ?? event.target.dataset.groupRoomRooms);
  if (!Number.isInteger(idx) || !state.groupDraft.roomItems[idx]) return;
  const item = state.groupDraft.roomItems[idx];
  let needsFullRender = false;
  if (event.target.dataset.groupRoomSubgroup !== undefined) item.subgroup = clean(event.target.value);
  if (event.target.dataset.groupRoomType !== undefined) {
    item.roomType = clean(event.target.value);
    item.rooms = [];
    needsFullRender = true;
  }
  if (event.target.dataset.groupRoomCount !== undefined) {
    item.roomCount = normalizeGroupRoomCount(event.target.value);
    item.rooms = (item.rooms || []).slice(0, item.roomCount);
    needsFullRender = true;
  }
  if (event.target.dataset.groupRoomPrice !== undefined) item.price = normalizeNumber(event.target.value);
  if (event.target.dataset.groupRoomUnder13 !== undefined) item.under13 = event.target.checked;
  if (event.target.dataset.groupRoomMode !== undefined) item.priceMode = clean(event.target.value) || "guest";
  if (event.target.dataset.groupRoomRooms !== undefined) {
    item.rooms = normalizeSelectedGroupRooms(idx, Array.from(event.target.selectedOptions).map((option) => clean(option.value)));
    needsFullRender = true;
  }
  if (needsFullRender) renderGroupDraft();
  else renderGroupTotals();
}

function onGroupRoomItemAction(event) {
  const btn = event.target.closest("[data-remove-group-room]");
  if (!btn) return;
  state.groupDraft.roomItems.splice(Number(btn.dataset.removeGroupRoom), 1);
  renderGroupDraft();
}

async function onGroupRowClick(event) {
  const row = event.target.closest("[data-group-id]");
  if (!row) return;
  await refreshGroupSettingsForEditor();
  await loadGroups({ silent: true });
  state.groupsLoaded = true;
  const group = state.groups.find((item) => item.id === clean(row.dataset.groupId));
  if (!group) {
    renderGroups();
    showToast("This group proposal is no longer available.", "error");
    return;
  }
  state.groupSelectedId = group.id;
  state.groupDraft = clone(group);
  openGroupModal();
  renderGroups();
}

async function refreshGroupSettingsForEditor() {
  if (!canApp("groups") && !canSettings("groups")) return;
  await loadGroupSettings({ silent: true });
}

async function saveGroupProposal() {
  const previous = state.groups.find((item) => item.id === clean(state.groupDraft.id));
  const payload = groupDraftPayload(previous);
  if (!payload) return;
  try {
    const id = clean(state.groupDraft.id);
    await api(id ? `/api/groups?id=${encodeURIComponent(id)}` : "/api/groups", {
      method: id ? "PUT" : "POST",
      body: payload,
    });
    state.groupsLoaded = false;
    await loadGroups();
    closeGroupModal();
    resetGroupDraft();
    setGroupsStatus("Group proposal saved.");
    showToast("Group proposal saved.", "success");
  } catch (e) {
    setGroupsStatus(`Could not save group proposal: ${e.message}`);
    showToast(`Could not save group proposal: ${e.message}`, "error");
  }
}

async function deleteGroupProposal() {
  if (!isAdministratorProfile()) {
    setGroupsStatus("Only Administrator can delete group proposals.");
    showToast("Only Administrator can delete group proposals.", "error");
    return;
  }
  const id = clean(state.groupDraft.id);
  if (!id || !window.confirm("Delete this group proposal?")) return;
  try {
    await api(`/api/groups?id=${encodeURIComponent(id)}`, { method: "DELETE" });
    state.groupsLoaded = false;
    await loadGroups();
    resetGroupDraft();
    closeGroupModal();
    setGroupsStatus("Group proposal deleted.");
  } catch (e) {
    setGroupsStatus(`Could not delete group proposal: ${e.message}`);
  }
}

function groupDraftPayload(previous = null) {
  const draft = state.groupDraft;
  if (!draft.name) return setGroupsStatus("Name is required."), null;
  if (!draft.email || !draft.email.includes("@")) return setGroupsStatus("A valid email is required."), null;
  if (normalizeGroupStatus(draft.status) === "Accepted" && !clean(draft.reservationNumber)) {
    return setGroupsStatus("Reservation Number is required when the proposal is Accepted."), null;
  }
  if (!draft.checkIn || !draft.checkOut) return setGroupsStatus("Check-in and check-out are required."), null;
  if (draft.checkOut <= draft.checkIn) return setGroupsStatus("Check-out must be after check-in."), null;
  const guests = normalizeGroupGuests(draft.guests);
  if (!guests || guests > 60) return setGroupsStatus("Guests must be between 1 and 60."), null;
  const duplicateRoom = firstDuplicateGroupRoom(draft.roomItems);
  if (duplicateRoom) return setGroupsStatus(`Room ${duplicateRoom} is selected more than once in this proposal.`), null;
  const roomItems = draft.roomItems
    .filter((item) => clean(item.roomType))
    .map((item) => ({
      subgroup: clean(item.subgroup),
      roomType: clean(item.roomType),
      roomCount: normalizeGroupRoomCount(item.roomCount),
      guests: normalizeGroupGuests(item.guests, 0),
      under13: !!item.under13,
      price: normalizeNumber(item.price),
      priceMode: clean(item.priceMode) === "room" ? "room" : "guest",
      rooms: normalizeSelectedGroupRooms(-1, item.rooms || []).slice(0, normalizeGroupRoomCount(item.roomCount)),
      lineTotal: calculateGroupRoomItemTotal(draft, item),
    }));
  const audit = appendGroupAudit(draft, previous);
  draft.audit = audit;
  return {
    reservationNumber: draft.reservationNumber,
    name: draft.name,
    email: draft.email,
    checkIn: draft.checkIn,
    checkOut: draft.checkOut,
    guests,
    guestGroups: buildGroupMetadata(draft, audit),
    roomItems,
    totalValue: calculateGroupTotal(draft),
    optionDate: draft.optionDate || null,
    status: normalizeGroupStatus(draft.status),
  };
}

function appendGroupAudit(draft, previous) {
  const currentAudit = normalizeGroupAudit(draft.audit);
  const summary = groupAuditSummary(draft, previous);
  const action = previous?.id ? "Updated proposal" : "Created proposal";
  return currentAudit.concat([{ at: new Date().toISOString(), action, summary }]).slice(-20);
}

function groupAuditSummary(draft, previous) {
  if (!previous?.id) return `Created for ${clean(draft.name) || "group"} (${formatGroupDateDisplay(draft.checkIn)} - ${formatGroupDateDisplay(draft.checkOut)}), total ${formatMoney(calculateGroupTotal(draft))}.`;
  const changes = [];
  if (clean(previous.status) !== clean(draft.status)) changes.push(`status ${previous.status || "-"} -> ${draft.status || "-"}`);
  if (clean(previous.reservationNumber) !== clean(draft.reservationNumber)) changes.push("reservation number changed");
  if (clean(previous.checkIn) !== clean(draft.checkIn) || clean(previous.checkOut) !== clean(draft.checkOut)) changes.push("dates changed");
  if (Number(previous.guests || 0) !== Number(draft.guests || 0)) changes.push(`guests ${previous.guests || 0} -> ${draft.guests || 0}`);
  if (Math.abs(Number(previous.totalValue || 0) - calculateGroupTotal(draft)) >= 0.01) changes.push(`total ${formatMoney(previous.totalValue || 0)} -> ${formatMoney(calculateGroupTotal(draft))}`);
  if (normalizeProposalLanguage(previous.language) !== normalizeProposalLanguage(draft.language || state.groupProposalLanguage)) changes.push(`language ${proposalLanguageLabel(previous.language)} -> ${proposalLanguageLabel(draft.language || state.groupProposalLanguage)}`);
  if (clean(previous.observation) !== clean(draft.observation)) changes.push("observation changed");
  if (JSON.stringify(previous.roomItems || []) !== JSON.stringify(draft.roomItems || [])) changes.push("room lines changed");
  return changes.length ? changes.join("; ") : "Saved without major field changes.";
}

function resetGroupDraft() {
  state.groupSelectedId = "";
  state.groupDraft = emptyGroupDraft();
  renderGroups();
}

function openGroupModal() {
  state.groupEditorTab = "details";
  els.groupEditorModal.hidden = false;
  document.body.classList.add("modal-open");
  renderGroupDraft();
}

function closeGroupModal() {
  els.groupEditorModal.hidden = true;
  document.body.classList.remove("modal-open");
}

function calculateGroupTotal(draft) {
  return calculateGroupAccommodationTotal(draft) + calculateGroupCityTaxTotal(draft);
}

function calculateGroupAccommodationTotal(draft) {
  syncGroupRoomItemGuests(draft);
  return (draft.roomItems || []).reduce((total, item) => total + calculateGroupRoomItemTotal(draft, item), 0);
}

function calculateGroupRoomItemTotal(draft, item) {
  const nights = Math.max(1, dateDiffDays(draft.checkIn, draft.checkOut));
  const price = normalizeNumber(item.price);
  const roomCount = normalizeGroupRoomCount(item.roomCount);
  const rowGuests = normalizeGroupGuests(item.guests, 0);
  const quantity = item.priceMode === "room" ? roomCount : rowGuests;
  return price * quantity * nights;
}

function calculateGroupCityTaxTotal(draft) {
  const taxableNights = groupCityTaxableNights(draft);
  return groupCityTaxableGuests(draft) * taxableNights * 4;
}

function groupCityTaxableNights(draft) {
  return Math.min(7, Math.max(1, dateDiffDays(draft.checkIn, draft.checkOut)));
}

function groupCityTaxableGuests(draft) {
  syncGroupRoomItemGuests(draft);
  const items = (draft.roomItems || []).filter((item) => clean(item.roomType));
  if (!items.length) return normalizeGroupGuests(draft.guests, 0);
  const assignedGuests = items.reduce((sum, item) => sum + normalizeGroupGuests(item.guests, 0), 0);
  const taxableAssignedGuests = items.reduce((sum, item) => sum + (item.under13 ? 0 : normalizeGroupGuests(item.guests, 0)), 0);
  const unassignedGuests = Math.max(0, normalizeGroupGuests(draft.guests, 0) - assignedGuests);
  return taxableAssignedGuests + unassignedGuests;
}

function syncGroupRoomItemGuests(draft) {
  (draft.roomItems || []).forEach((item) => {
    const roomCount = normalizeGroupRoomCount(item.roomCount);
    const guestsPerRoom = guestsPerGroupRoomType(item.roomType);
    item.roomCount = roomCount;
    item.guests = clean(item.roomType) ? roomCount * guestsPerRoom : 0;
  });
}

function normalizeGroupRoomItems(items) {
  return (Array.isArray(items) ? items : []).map((item) => {
    const roomType = clean(item.roomType);
    const guests = normalizeGroupGuests(item.guests, 0);
    const configuredGuestsPerRoom = guestsPerGroupRoomType(roomType);
    const guestsPerRoom = configuredGuestsPerRoom || 1;
    const roomCount = item.roomCount ? normalizeGroupRoomCount(item.roomCount) : Math.max(1, Math.ceil(guests / guestsPerRoom));
    return {
      subgroup: clean(item.subgroup),
      roomType,
      roomCount,
      guests: configuredGuestsPerRoom ? roomCount * configuredGuestsPerRoom : guests,
      under13: !!item.under13,
      price: normalizeNumber(item.price),
      priceMode: clean(item.priceMode) === "room" ? "room" : "guest",
      rooms: Array.isArray(item.rooms) ? item.rooms.map(clean).filter(Boolean) : [],
    };
  });
}

function groupRemainingGuests(draft) {
  syncGroupRoomItemGuests(draft);
  const totalGuests = normalizeGroupGuests(draft.guests, 0);
  const assigned = (draft.roomItems || []).reduce((sum, item) => sum + normalizeGroupGuests(item.guests, 0), 0);
  return totalGuests - assigned;
}

function dateDiffDays(start, end) {
  const a = new Date(start);
  const b = new Date(end);
  if (Number.isNaN(a.getTime()) || Number.isNaN(b.getTime())) return 1;
  return Math.max(1, Math.round((b.getTime() - a.getTime()) / (24 * 60 * 60 * 1000)));
}

function formatGroupDateInput(value) {
  const raw = clean(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  return raw;
}

function formatGroupDateDisplay(value) {
  const raw = clean(value);
  return raw ? formatGroupDateInput(raw) : "-";
}

function parseGroupDateInput(value) {
  const raw = clean(value);
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const isoLike = raw.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (isoLike) {
    const year = isoLike[1];
    const month = isoLike[2].padStart(2, "0");
    const day = isoLike[3].padStart(2, "0");
    const iso = `${year}-${month}-${day}`;
    const dt = new Date(iso);
    if (Number.isNaN(dt.getTime())) return "";
    return formatDate(dt) === iso ? iso : "";
  }
  const match = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (!match) return "";
  const day = match[1].padStart(2, "0");
  const month = match[2].padStart(2, "0");
  const year = match[3];
  const iso = `${year}-${month}-${day}`;
  const dt = new Date(iso);
  if (Number.isNaN(dt.getTime())) return "";
  return formatDate(dt) === iso ? iso : "";
}

function groupNightsLabelText(checkIn, checkOut) {
  if (!clean(checkIn) || !clean(checkOut)) return "Nights: -";
  if (checkOut <= checkIn) return "Nights: -";
  const nights = dateDiffDays(checkIn, checkOut);
  return `Nights: ${nights}`;
}

function groupLastPaymentLimitText(checkIn) {
  const raw = clean(checkIn);
  if (!raw) return "-";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return "-";
  date.setDate(date.getDate() - normalizeLastPaymentDays(state.groupSettings.lastPaymentDaysBeforeArrival));
  return formatGroupDateDisplay(formatDate(date));
}

function normalizeLastPaymentDays(value) {
  const num = Math.round(normalizeNumber(value, 14));
  return Math.max(0, Math.min(365, num));
}

function renderGroupStatusColor() {
  const status = normalizeGroupStatus(els.groupStatusField.value);
  els.groupStatusField.classList.toggle("status-accepted", status === "Accepted");
  els.groupStatusField.classList.toggle("status-refused", status === "Refused");
  els.groupStatusField.classList.toggle("status-proposal", status === "Proposal");
}

function roomsForGroupRoomType(name) {
  return groupRoomTypeConfig(name)?.rooms || [];
}

function selectedGroupRoomsExcept(index) {
  const rooms = new Set();
  (state.groupDraft.roomItems || []).forEach((item, itemIndex) => {
    if (itemIndex === index) return;
    (Array.isArray(item.rooms) ? item.rooms : []).map(clean).filter(Boolean).forEach((room) => rooms.add(room));
  });
  return rooms;
}

function normalizeSelectedGroupRooms(index, rooms) {
  const usedElsewhere = index >= 0 ? selectedGroupRoomsExcept(index) : new Set();
  const unique = [];
  (Array.isArray(rooms) ? rooms : []).map(clean).filter(Boolean).forEach((room) => {
    if (usedElsewhere.has(room) || unique.includes(room)) return;
    unique.push(room);
  });
  const item = state.groupDraft.roomItems[index];
  const roomLimit = item ? normalizeGroupRoomCount(item.roomCount) : 20;
  return unique.slice(0, roomLimit);
}

function firstDuplicateGroupRoom(items = []) {
  const seen = new Set();
  for (const item of items) {
    for (const room of Array.isArray(item.rooms) ? item.rooms.map(clean).filter(Boolean) : []) {
      if (seen.has(room)) return room;
      seen.add(room);
    }
  }
  return "";
}

function guestsPerGroupRoomType(name) {
  if (!clean(name)) return 0;
  return groupRoomTypeConfig(name)?.guestsPerRoom || inferGuestsPerGroupRoomType(name);
}

function groupRoomTypeConfig(name) {
  return state.groupSettings.roomTypes.find((item) => item.name === name);
}

function inferGuestsPerGroupRoomType(name) {
  const raw = clean(name);
  const leadingNumber = Number(raw.match(/^\d+/)?.[0]);
  if (Number.isFinite(leadingNumber) && leadingNumber > 0) return Math.min(20, leadingNumber);
  if (/twin/i.test(raw)) return 2;
  if (/single/i.test(raw)) return 1;
  return 1;
}

function normalizeGroupRoomCount(value) {
  const num = Math.round(normalizeNumber(value, 1));
  return Math.max(1, Math.min(20, num));
}

function normalizeGroupGuests(value, fallback = 1) {
  const num = Math.round(normalizeNumber(value, fallback));
  return Math.max(0, Math.min(60, num));
}

function normalizeGroupStatus(value) {
  const raw = clean(value).toLowerCase();
  if (raw === "accepted") return "Accepted";
  if (raw === "refused") return "Refused";
  return "Proposal";
}

function sanitizeGroupSettings(settings) {
  const source = settings && typeof settings === "object" ? settings : {};
  const roomTypes = Array.isArray(source.roomTypes) ? source.roomTypes : [];
  return {
    depositPercentage: Math.max(0, Math.min(100, normalizeNumber(source.depositPercentage, 30))),
    lastPaymentDaysBeforeArrival: normalizeLastPaymentDays(source.lastPaymentDaysBeforeArrival ?? source.last_payment_days_before_arrival),
    emailTemplate: clean(source.emailTemplate) || DEFAULT_GROUP_SETTINGS.emailTemplate,
    confirmationTemplate: clean(source.confirmationTemplate) || DEFAULT_GROUP_SETTINGS.confirmationTemplate,
    finalConfirmationTemplate: clean(source.finalConfirmationTemplate) || DEFAULT_GROUP_SETTINGS.finalConfirmationTemplate,
    roomTypes: roomTypes.length ? roomTypes.map((item) => ({
      name: clean(item.name),
      guestsPerRoom: Math.max(1, Math.min(20, Math.round(normalizeNumber(item.guestsPerRoom ?? item.guests_per_room, inferGuestsPerGroupRoomType(item.name))))),
      rooms: Array.isArray(item.rooms) ? item.rooms.map(clean).filter(Boolean) : clean(item.rooms).split(",").map(clean).filter(Boolean),
    })).filter((item) => item.name) : clone(DEFAULT_GROUP_ROOM_TYPES),
  };
}

function renderGroupSettings() {
  if (!els.groupsRoomTypesBody) return;
  renderGroupSettingsTab();
  els.groupsDepositPercentage.value = state.groupSettings.depositPercentage;
  els.groupsLastPaymentDays.value = normalizeLastPaymentDays(state.groupSettings.lastPaymentDaysBeforeArrival);
  els.groupsEmailTemplate.value = state.groupSettings.emailTemplate || DEFAULT_GROUP_SETTINGS.emailTemplate;
  els.groupsConfirmationTemplate.value = state.groupSettings.confirmationTemplate || DEFAULT_GROUP_SETTINGS.confirmationTemplate;
  els.groupsFinalConfirmationTemplate.value = state.groupSettings.finalConfirmationTemplate || DEFAULT_GROUP_SETTINGS.finalConfirmationTemplate;
  renderGroupSettingsTemplatePreviews();
  els.groupsRoomTypesBody.innerHTML = "";
  state.groupSettings.roomTypes.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input data-group-setting-room-name="${index}" value="${escape(item.name)}" /></td>
      <td><input class="small-number-input" type="number" min="1" max="20" step="1" data-group-setting-room-guests="${index}" value="${escape(item.guestsPerRoom || inferGuestsPerGroupRoomType(item.name))}" /></td>
      <td><input data-group-setting-room-rooms="${index}" value="${escape((item.rooms || []).join(", "))}" /></td>
      <td><button type="button" class="danger" data-remove-group-setting-room="${index}">Remove</button></td>`;
    els.groupsRoomTypesBody.appendChild(tr);
  });
}

function renderGroupSettingsTab() {
  const isConfig = state.groupSettingsTab === "config";
  const isProposal = state.groupSettingsTab === "proposal";
  const isConfirmation = state.groupSettingsTab === "confirmation";
  const isFinalConfirmation = state.groupSettingsTab === "final-confirmation";
  els.groupsSettingsConfigTab.classList.toggle("active-tab", isConfig);
  els.groupsSettingsConfigTab.classList.toggle("ghost", !isConfig);
  els.groupsSettingsProposalTab.classList.toggle("active-tab", isProposal);
  els.groupsSettingsProposalTab.classList.toggle("ghost", !isProposal);
  els.groupsSettingsConfirmationTab.classList.toggle("active-tab", isConfirmation);
  els.groupsSettingsConfirmationTab.classList.toggle("ghost", !isConfirmation);
  els.groupsSettingsFinalConfirmationTab.classList.toggle("active-tab", isFinalConfirmation);
  els.groupsSettingsFinalConfirmationTab.classList.toggle("ghost", !isFinalConfirmation);
  els.groupsSettingsConfigPanel.hidden = !isConfig;
  els.groupsSettingsProposalPanel.hidden = !isProposal;
  els.groupsSettingsConfirmationPanel.hidden = !isConfirmation;
  els.groupsSettingsFinalConfirmationPanel.hidden = !isFinalConfirmation;
}

function renderGroupSettingsTemplatePreviews() {
  const draft = groupTemplatePreviewDraft();
  if (els.groupsProposalTemplatePreview) {
    els.groupsProposalTemplatePreview.innerHTML = `<h4>Preview</h4>${groupTemplatePreviewHtml(draft, clean(state.groupSettings.emailTemplate) || DEFAULT_GROUP_SETTINGS.emailTemplate, "proposal")}`;
  }
  if (els.groupsConfirmationTemplatePreview) {
    els.groupsConfirmationTemplatePreview.innerHTML = `<h4>Preview</h4>${groupTemplatePreviewHtml(draft, clean(state.groupSettings.confirmationTemplate) || DEFAULT_GROUP_SETTINGS.confirmationTemplate, "first")}`;
  }
  if (els.groupsFinalConfirmationTemplatePreview) {
    els.groupsFinalConfirmationTemplatePreview.innerHTML = `<h4>Preview</h4>${groupTemplatePreviewHtml(draft, clean(state.groupSettings.finalConfirmationTemplate) || DEFAULT_GROUP_SETTINGS.finalConfirmationTemplate, "final")}`;
  }
}

function groupTemplatePreviewDraft() {
  const roomType = state.groupSettings.roomTypes[0]?.name || "10 Bed Dorm Shared Bathroom";
  return {
    ...emptyGroupDraft(),
    reservationNumber: "123456",
    name: "Sample Group",
    email: "group@example.com",
    checkIn: "2026-05-12",
    checkOut: "2026-05-15",
    guests: 10,
    status: "Accepted",
    language: "en",
    roomItems: [{ subgroup: "students", roomType, roomCount: 1, guests: guestsPerGroupRoomType(roomType), under13: false, price: 30, priceMode: "guest", rooms: [] }],
  };
}

function groupTemplatePreviewHtml(draft, template, kind) {
  const previousLanguage = state.groupProposalLanguage;
  state.groupProposalLanguage = "en";
  try {
    if (kind === "proposal") return groupProposalEmailHtmlFromTemplate(draft, template);
    return groupConfirmationEmailHtmlFromTemplate(draft, template, kind);
  } finally {
    state.groupProposalLanguage = previousLanguage;
  }
}

function setGroupSettingsTab(tab) {
  state.groupSettingsTab = ["proposal", "confirmation", "final-confirmation"].includes(tab) ? tab : "config";
  renderGroupSettingsTab();
}

function onGroupSettingsInput(event) {
  state.groupSettings.depositPercentage = Math.max(0, Math.min(100, normalizeNumber(els.groupsDepositPercentage.value, 0)));
  state.groupSettings.lastPaymentDaysBeforeArrival = normalizeLastPaymentDays(els.groupsLastPaymentDays.value);
  state.groupSettings.emailTemplate = els.groupsEmailTemplate.value;
  state.groupSettings.confirmationTemplate = els.groupsConfirmationTemplate.value;
  state.groupSettings.finalConfirmationTemplate = els.groupsFinalConfirmationTemplate.value;
  renderGroupSettingsTemplatePreviews();
  const nameIdx = event.target.dataset.groupSettingRoomName;
  const guestsIdx = event.target.dataset.groupSettingRoomGuests;
  const roomsIdx = event.target.dataset.groupSettingRoomRooms;
  const idx = Number(nameIdx ?? guestsIdx ?? roomsIdx);
  if (Number.isInteger(idx) && state.groupSettings.roomTypes[idx]) {
    if (nameIdx !== undefined) state.groupSettings.roomTypes[idx].name = clean(event.target.value);
    if (guestsIdx !== undefined) state.groupSettings.roomTypes[idx].guestsPerRoom = Math.max(1, Math.min(20, Math.round(normalizeNumber(event.target.value, 1))));
    if (roomsIdx !== undefined) state.groupSettings.roomTypes[idx].rooms = clean(event.target.value).split(",").map(clean).filter(Boolean);
  }
  if (!els.groupEditorModal.hidden) renderGroupDraft();
}

function addGroupSettingsRoomType() {
  state.groupSettings.roomTypes.push({ name: "New Room Type", guestsPerRoom: 1, rooms: [] });
  renderGroupSettings();
}

function onGroupSettingsRoomTypeAction(event) {
  const btn = event.target.closest("[data-remove-group-setting-room]");
  if (!btn) return;
  state.groupSettings.roomTypes.splice(Number(btn.dataset.removeGroupSettingRoom), 1);
  renderGroupSettings();
}

async function saveGroupSettings() {
  try {
    state.groupSettings = sanitizeGroupSettings(state.groupSettings);
    await api("/api/group-settings", { method: "PUT", body: { settings: state.groupSettings } });
    renderGroupSettings();
    if (!els.groupEditorModal.hidden) renderGroupDraft();
    setGroupsSettingsStatus("Group settings saved.");
    showToast("Group settings saved.", "success");
  } catch (e) {
    setGroupsSettingsStatus(`Could not save group settings: ${e.message}`);
  }
}

async function loadEntries({ silent = false } = {}) {
  try {
    const result = await api("/api/communications");
    state.entries = (result.rows || []).map((row) => ({
      id: row.id,
      date: normalizeDate(clean(row.date)),
      time: normalizeTime(clean(row.time)),
      person: clean(row.person),
      status: normalizeStatusUi(row.status),
      category: normalizeCategory(row.category),
      message: clean(row.message),
      createdAt: clean(row.created_at),
      updatedAt: clean(row.updated_at),
    }));
    if (!silent) setDbStatus(`Loaded ${state.entries.length} records.`);
  } catch (e) {
    setDbStatus(`DB error: ${e.message}`);
    showToast(`DB error: ${e.message}`, "error");
  }
}

function setLostFoundStatus(message) {
  if (els.lostFoundDbStatus) els.lostFoundDbStatus.textContent = message;
}

function lostFoundDisplayNumber(value) {
  const num = Number(value) || 0;
  return num > 0 ? num + LOST_FOUND_NUMBER_OFFSET : 0;
}

async function loadLostFound({ silent = false } = {}) {
  try {
    const result = await api("/api/lost-found");
    state.lostFound = (result.rows || []).map((row) => ({
      id: row.id,
      number: lostFoundDisplayNumber(row.item_number),
      createdAt: clean(row.created_at),
      updatedAt: clean(row.updated_at),
      closedAt: clean(row.closed_at),
      whoFound: clean(row.who_found),
      whoRecorded: clean(row.who_recorded),
      location: clean(row.location_found),
      objectDescription: clean(row.object_description),
      notes: clean(row.notes),
      stored: normalizeLostFoundStored(clean(row.stored_location)),
      status: normalizeStatusUi(row.status),
    }));
    if (!silent) setLostFoundStatus(`Loaded ${state.lostFound.length} records.`);
  } catch (error) {
    setLostFoundStatus(`DB error: ${error.message}`);
    showToast(`DB error: ${error.message}`, "error");
  }
}

function renderSettings() {
  const general = state.settings.general?.emailConfig || normalizeBakeryEmailConfigClient();
  if (els.generalEmailProvider) els.generalEmailProvider.value = general.provider;
  if (els.generalEmailSmtpHost) els.generalEmailSmtpHost.value = general.smtpHost;
  if (els.generalEmailSmtpPort) els.generalEmailSmtpPort.value = general.smtpPort;
  if (els.generalEmailSmtpSecure) els.generalEmailSmtpSecure.checked = !!general.smtpSecure;
  if (els.generalEmailSmtpUser) els.generalEmailSmtpUser.value = general.smtpUser;
  if (els.generalEmailSmtpPassword) els.generalEmailSmtpPassword.value = general.smtpPassword;
  if (els.generalEmailFromEmail) els.generalEmailFromEmail.value = general.fromEmail;
  if (els.generalEmailFromName) els.generalEmailFromName.value = general.fromName;
  if (els.generalEmailSmtpFields) els.generalEmailSmtpFields.hidden = general.provider !== "smtp";
  const cfg = state.settings.communications;
  els.settingsCategoriesBody.innerHTML = "";
  cfg.categories.forEach((cat, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input data-idx="${index}" data-field="name" value="${escape(cat.name)}" /></td>
      <td><input type="color" data-idx="${index}" data-field="color" value="${escape(cat.color)}" /></td>
      <td><input type="number" min="1" step="1" data-idx="${index}" data-field="autoCloseDays" value="${escape(cat.autoCloseDays ?? "")}" placeholder="Manual" /></td>
      <td><span class="chip" style="${chipStyle(cat.color)}">${escape(cat.name)}</span></td>
      <td><button class="danger" type="button" data-remove="${index}">Remove</button></td>`;
    els.settingsCategoriesBody.appendChild(tr);
  });
  els.settingEmailEnabled.checked = !!cfg.emailAutomation.enabled;
  els.settingEmailFrequency.value = cfg.emailAutomation.frequency;
  els.settingEmailTime.value = cfg.emailAutomation.timeOfDay;
  els.settingEmailRecipients.value = (cfg.emailAutomation.recipients || []).join("\n");
  els.settingEmailFrequency2.value = cfg.emailAutomation.frequency2 || "everyday";
  els.settingEmailTime2.value = cfg.emailAutomation.timeOfDay2 || "00:00";
  els.settingEmailRecipients2.value = (cfg.emailAutomation.recipients2 || []).join("\n");
  els.settingEmailPreview.textContent = emailPreview(cfg.emailAutomation);
  els.settingEmailNextPreview.textContent = nextSendTimesPreview(cfg.emailAutomation);
}

function onGeneralSettingsInput(event) {
  const general = state.settings.general.emailConfig;
  if (event.target === els.generalEmailProvider) general.provider = normalizeBakeryEmailProviderClient(event.target.value);
  if (event.target === els.generalEmailSmtpHost) general.smtpHost = clean(event.target.value);
  if (event.target === els.generalEmailSmtpPort) general.smtpPort = Math.max(1, Number.parseInt(event.target.value, 10) || 1);
  if (event.target === els.generalEmailSmtpSecure) general.smtpSecure = !!event.target.checked;
  if (event.target === els.generalEmailSmtpUser) general.smtpUser = clean(event.target.value).toLowerCase();
  if (event.target === els.generalEmailSmtpPassword) general.smtpPassword = String(event.target.value || "");
  if (event.target === els.generalEmailFromEmail) general.fromEmail = clean(event.target.value).toLowerCase();
  if (event.target === els.generalEmailFromName) general.fromName = clean(event.target.value);
  state.settings = sanitizeSettings(state.settings);
  renderSettings();
}

function addCategory() {
  const list = state.settings.communications.categories;
  let i = 1;
  let name = "New Category";
  const set = new Set(list.map((x) => x.name.toLowerCase()));
  while (set.has(name.toLowerCase())) {
    i += 1;
    name = `New Category ${i}`;
  }
  list.push({ name, color: "#d8d8d8", autoCloseDays: null });
  renderSettings();
}

function removeCategoryClick(event) {
  const btn = event.target.closest("button[data-remove]");
  if (!btn) return;
  const idx = Number(btn.dataset.remove);
  const list = state.settings.communications.categories;
  if (list.length <= 1) return setSettingsStatus("At least one category is required.");
  list.splice(idx, 1);
  normalizeDraftsToSettings();
  renderCategoryFilterOptions();
  renderSettings();
  render();
}

function settingsCategoryInput(event) {
  const target = event.target;
  const idx = Number(target.dataset.idx);
  const field = target.dataset.field;
  if (Number.isNaN(idx) || !field) return;
  const item = state.settings.communications.categories[idx];
  if (!item) return;
  if (field === "color") item[field] = normalizeHex(target.value);
  else if (field === "autoCloseDays") item[field] = normalizeAutoCloseDays(target.value);
  else item[field] = clean(target.value);
  state.settings = sanitizeSettings(state.settings);
  normalizeDraftsToSettings();
  renderCategoryFilterOptions();
  renderSettings();
  render();
}

function updateEmailSettings() {
  const email = state.settings.communications.emailAutomation;
  email.enabled = els.settingEmailEnabled.checked;
  email.frequency = normalizeFrequency(els.settingEmailFrequency.value);
  email.timeOfDay = normalizeTimeInput(els.settingEmailTime.value);
  email.recipients = parseEmailList(els.settingEmailRecipients.value);
  email.frequency2 = normalizeFrequency(els.settingEmailFrequency2.value);
  email.timeOfDay2 = normalizeTimeInput(els.settingEmailTime2.value);
  email.recipients2 = parseEmailList(els.settingEmailRecipients2.value);
  els.settingEmailPreview.textContent = emailPreview(email);
  els.settingEmailNextPreview.textContent = nextSendTimesPreview(email);
}

function emailPreview(email) {
  if (!email.enabled) return "Automatic emails are disabled.";
  const lines = emailScheduleSummaries(email);
  if (!lines.length) return "Automatic emails are enabled but no recipient emails are configured.";
  return lines.join(" ");
}

function nextSendTimesPreview(email) {
  if (!email.enabled) return "Next send times preview: automatic emails are disabled.";
  const schedules = communicationEmailSchedules(email).filter((schedule) => schedule.recipients.length);
  if (!schedules.length) return "Next send times preview: add at least one recipient email.";
  const parts = schedules.map((schedule) => {
    const list = computeNextSendTimes(schedule, 5);
    return `${schedule.label}: ${list.length ? list.join(" | ") : "unable to compute"}`;
  });
  return `Next send times preview: ${parts.join(" || ")}`;
}

function computeNextSendTimes(email, count = 5) {
  const [hRaw, mRaw] = normalizeTimeInput(email.timeOfDay).split(":");
  const startHour = Number(hRaw);
  const startMinute = Number(mRaw);
  if (Number.isNaN(startHour) || Number.isNaN(startMinute)) return [];
  const now = new Date();
  const candidates = [];
  const daily = email.frequency === "everyday";
  const step = emailFrequencyStep(email.frequency);
  for (let dayOffset = 0; dayOffset <= 14 && candidates.length < count; dayOffset += 1) {
    const base = new Date(now);
    base.setHours(0, 0, 0, 0);
    base.setDate(base.getDate() + dayOffset);
    if (daily) {
      const dt = new Date(base);
      dt.setHours(startHour, startMinute, 0, 0);
      if (dt > now) candidates.push(dt);
      continue;
    }
    for (let hour = startHour; hour < 24; hour += step) {
      const dt = new Date(base);
      dt.setHours(hour, startMinute, 0, 0);
      if (dt > now) candidates.push(dt);
      if (candidates.length >= count) break;
    }
  }
  return candidates.slice(0, count).map(formatDateTimeShort);
}

async function triggerEmailNow() {
  const testRecipient = clean(els.settingEmailTestRecipient.value).toLowerCase();
  if (!isValidEmail(testRecipient)) {
    return setSettingsStatus("Please provide a valid test recipient email.");
  }
  els.testEmailNow.disabled = true;
  setSettingsStatus("Sending test email...");
  try {
    await api("/api/email-automation?force=1", {
      method: "POST",
      body: { testRecipient },
    });
    setSettingsStatus(`Test email sent successfully to ${testRecipient}.`);
  } catch (e) {
    setSettingsStatus(`Test email failed: ${e.message}`);
  } finally {
    els.testEmailNow.disabled = false;
  }
}

function renderCategoryFilterOptions() {
  const old = els.categoryFilter.value;
  const items = getCategories();
  els.categoryFilter.innerHTML = `<option value="">All</option>${items
    .map((x) => `<option value="${escape(x.name)}">${escape(x.name)}</option>`)
    .join("")}`;
  els.categoryFilter.value = items.some((x) => x.name === old) ? old : "";
}

function onRowDraftInput(event) {
  const t = event.target;
  const field = t.dataset.field;
  const scope = t.dataset.scope;
  if (!field || !scope) return;
  const value = field === "status" ? (t.checked ? "Closed" : "Open") : clean(t.value);
  if (scope === "new") {
    state.newDraft[field] = value;
    const row = t.closest("tr, .communication-mobile-card");
    if (row?.style) row.style.backgroundColor = "#ffffff";
  }
  if (scope === "edit" && state.editingId && t.dataset.id === state.editingId) {
    state.editDraft[field] = value;
    const row = t.closest("tr, .communication-mobile-card");
    if (row?.style) row.style.backgroundColor = rowBackgroundColor(state.editDraft.status, state.editDraft.category);
  }
}

async function triggerLaundryEmailNow() {
  els.laundryTestEmail.disabled = true;
  setLaundrySettingsStatus("Sending test email...");
  try {
    onLaundrySettingsInput();
    const saved = await api("/api/laundry-settings", { method: "PUT", body: { settings: state.laundrySettings } });
    state.laundrySettings = normalizeLaundrySettingsClient(saved?.settings);
    renderLaundrySettings();
    const result = await api("/api/laundry-email-automation?force=1", {
      method: "POST",
    });
    if (result?.status === "sent") {
      setLaundrySettingsStatus("Laundry test email sent successfully.");
    } else {
      setLaundrySettingsStatus(`Laundry test email was not sent: ${describeLaundryAutomationResult(result)}.`);
    }
  } catch (e) {
    setLaundrySettingsStatus(`Laundry test email failed: ${e.message}`);
  } finally {
    els.laundryTestEmail.disabled = false;
  }
}

async function triggerLaundryManagementEmailNow() {
  els.laundryManagementTestEmail.disabled = true;
  setLaundrySettingsStatus("Sending management test email...");
  try {
    onLaundrySettingsInput();
    const saved = await api("/api/laundry-settings", { method: "PUT", body: { settings: state.laundrySettings } });
    state.laundrySettings = normalizeLaundrySettingsClient(saved?.settings);
    renderLaundrySettings();
    const result = await api("/api/laundry-email-automation?force=1", {
      method: "POST",
      body: { mode: "management" },
    });
    if (result?.status === "sent") {
      setLaundrySettingsStatus("Laundry management test email sent successfully.");
    } else {
      setLaundrySettingsStatus(`Laundry management test email was not sent: ${describeLaundryAutomationResult(result)}.`);
    }
  } catch (e) {
    setLaundrySettingsStatus(`Laundry management test email failed: ${e.message}`);
  } finally {
    els.laundryManagementTestEmail.disabled = false;
  }
}

function describeLaundryAutomationResult(result = {}) {
  const reason = clean(result?.reason).toLowerCase();
  if (reason === "no_recipients") return "no recipient emails are configured";
  if (reason === "missing_records") {
    const missing = Array.isArray(result?.missingProperties) ? result.missingProperties.join(", ") : "required properties";
    return `records are missing for ${missing}`;
  }
  if (reason === "incomplete_counts") {
    const incomplete = Array.isArray(result?.incompleteProperties) ? result.incompleteProperties.join(", ") : "required properties";
    return `sent quantities are incomplete for ${incomplete}`;
  }
  if (reason === "disabled") return "the automation is disabled";
  if (reason === "time_mismatch") return "the configured time was not reached";
  if (reason === "already_sent_for_slot") return "it was already sent for this time slot";
  return reason || "the server skipped the send";
}

async function onRowAction(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) return;
  const action = button.dataset.action;
  if (action === "save-inline") return saveNew();
  if (action === "cancel-edit") {
    state.editingId = null;
    state.editDraft = null;
    showToast("Edit canceled.", "info");
    return render();
  }
  if (action === "save-edit") return saveEdit(button.dataset.id);
  const id = button.dataset.id;
  if (action === "edit") {
    await loadEntries({ silent: true });
    state.communicationsLoaded = true;
    const latestEntry = state.entries.find((x) => x.id === id);
    if (!latestEntry) {
      render();
      showToast("This communication is no longer available.", "error");
      return;
    }
    state.editingId = id;
    state.editDraft = { person: latestEntry.person, status: latestEntry.status, category: latestEntry.category, message: latestEntry.message };
    return render();
  }
  const entry = state.entries.find((x) => x.id === id);
  if (!entry) return;
  if (action === "delete") {
    button.disabled = true;
    const deletedIndex = state.entries.findIndex((x) => x.id === id);
    try {
      await api(`/api/communications?id=${encodeURIComponent(id)}`, { method: "DELETE" });
      if (deletedIndex !== -1) state.entries.splice(deletedIndex, 1);
      if (state.editingId === id) {
        state.editingId = null;
        state.editDraft = null;
      }
      queuePendingDelete(entry, deletedIndex);
      render();
      showToast("Record deleted.", "success", {
        actionLabel: "Undo",
        action: undoPendingDelete,
        duration: 9000,
      });
    } catch (error) {
      showToast(`Delete failed: ${error.message}`, "error");
    } finally {
      button.disabled = false;
    }
    return;
  }
}

async function onRowStatusToggle(event) {
  const input = event.target;
  if (!input.matches('input[data-action="toggle-status"]')) return;

  const id = clean(input.dataset.id);
  const entry = state.entries.find((x) => x.id === id);
  if (!entry) return;

  const nextStatus = input.checked ? "Closed" : "Open";
  input.disabled = true;
  try {
    await api(`/api/communications?id=${encodeURIComponent(id)}`, {
      method: "PUT",
      body: {
        person: entry.person,
        status: nextStatus,
        category: entry.category,
        message: entry.message,
      },
    });
    entry.status = nextStatus;
    entry.updatedAt = new Date().toISOString();
    render();
    showToast(`Status changed to ${nextStatus}.`, "success");
  } catch (error) {
    input.checked = !input.checked;
    showToast(`Status update failed: ${error.message}`, "error");
  } finally {
    input.disabled = false;
  }
}

async function saveNew() {
  const person = clean(state.newDraft.person);
  const message = clean(state.newDraft.message);
  if (!person || !message) return showToast("Please fill Person and What happened.", "error");
  const now = new Date();
  try {
    await api("/api/communications", {
      method: "POST",
      body: {
        date: formatDate(now),
        time: formatTime(now),
        person,
        status: normalizeStatusUi(state.newDraft.status),
        category: normalizeCategory(state.newDraft.category),
        message,
      },
    });
    state.newDraft = { person: "", status: "Open", category: getCategories()[0].name, message: "" };
    await loadEntries();
    render();
    showToast("Communication added.", "success");
  } catch (error) {
    showToast(`Save failed: ${error.message}`, "error");
  }
}

async function saveEdit(id) {
  const draft = state.editDraft;
  if (!id || !draft) return;
  if (!clean(draft.person) || !clean(draft.message)) return showToast("Please fill Person and What happened.", "error");
  try {
    await api(`/api/communications?id=${encodeURIComponent(id)}`, {
      method: "PUT",
      body: {
        person: clean(draft.person),
        status: normalizeStatusUi(draft.status),
        category: normalizeCategory(draft.category),
        message: clean(draft.message),
      },
    });
    state.editingId = null;
    state.editDraft = null;
    await loadEntries();
    render();
    showToast("Communication updated.", "success");
  } catch (error) {
    showToast(`Update failed: ${error.message}`, "error");
  }
}

function normalizeLostFoundStored(value) {
  const raw = clean(value).toLowerCase();
  return LOST_FOUND_STORED_OPTIONS.find((item) => item.toLowerCase() === raw) || LOST_FOUND_STORED_OPTIONS[0];
}

function emptyLostFoundDraft() {
  return {
    whoFound: "",
    whoRecorded: "",
    location: "",
    objectDescription: "",
    notes: "",
    stored: LOST_FOUND_STORED_OPTIONS[0],
    status: "Open",
  };
}

function lostFoundTimestampDate(record) {
  const dt = new Date(clean(record?.createdAt));
  return Number.isNaN(dt.getTime()) ? "" : formatDate(dt);
}

function lostFoundTimestampTime(record) {
  const dt = new Date(clean(record?.createdAt));
  return Number.isNaN(dt.getTime()) ? "" : formatTime(dt);
}

function nextLostFoundDisplayNumber() {
  return state.lostFound.reduce((max, record) => Math.max(max, Number(record.number) || 0), LOST_FOUND_NUMBER_OFFSET) + 1;
}

function lostFoundRowBackground(status) {
  return isClosedStatus(status) ? hexToRgba("#2e9f42", 0.25) : "#ffffff";
}

async function onLostFoundAction(event) {
  const button = event.target.closest("button[data-lost-found-action]");
  if (!button) return;
  const action = clean(button.dataset.lostFoundAction);
  if (action === "save-inline") return saveNewLostFound();
  if (action === "cancel-edit") {
    state.lostFoundEditingId = null;
    state.lostFoundEditDraft = null;
    showToast("Edit canceled.", "info");
    renderLostFound();
    return;
  }
  const id = clean(button.dataset.id);
  if (action === "edit") {
    await loadLostFound({ silent: true });
    state.lostFoundLoaded = true;
    const latestRecord = state.lostFound.find((item) => item.id === id);
    if (!latestRecord) {
      renderLostFound();
      showToast("This Lost&Found record is no longer available.", "error");
      return;
    }
    state.lostFoundEditingId = id;
    state.lostFoundEditDraft = {
      whoFound: latestRecord.whoFound,
      whoRecorded: latestRecord.whoRecorded,
      location: latestRecord.location,
      objectDescription: latestRecord.objectDescription,
      notes: latestRecord.notes,
      stored: latestRecord.stored,
      status: latestRecord.status,
    };
    renderLostFound();
    return;
  }
  const record = state.lostFound.find((item) => item.id === id);
  if (!record) return;
  if (action === "save-edit") {
    await saveLostFoundEdit(id);
  }
}

function onLostFoundDraftInput(event) {
  const target = event.target;
  const field = clean(target?.dataset?.field);
  const scope = clean(target?.dataset?.scope);
  if (!field || !scope) return;
  const value = field === "status"
    ? (target.checked ? "Closed" : "Open")
    : field === "stored"
      ? normalizeLostFoundStored(target.value)
      : clean(target.value);
  if (scope === "new") {
    state.lostFoundDraft[field] = value;
    const row = target.closest("tr, .lost-found-mobile-card");
    if (row) row.style.backgroundColor = "#ffffff";
  }
  if (scope === "edit" && state.lostFoundEditingId && clean(target.dataset.id) === state.lostFoundEditingId) {
    state.lostFoundEditDraft[field] = value;
    const row = target.closest("tr, .lost-found-mobile-card");
    if (row) row.style.backgroundColor = lostFoundRowBackground(state.lostFoundEditDraft.status);
  }
}

function onLostFoundKeydown(event) {
  const target = event.target;
  const scope = clean(target?.dataset?.scope);
  const id = clean(target?.dataset?.id);
  if (!scope) return;
  if (event.key === "Escape" && scope === "edit") {
    event.preventDefault();
    state.lostFoundEditingId = null;
    state.lostFoundEditDraft = null;
    showToast("Edit canceled.", "info");
    renderLostFound();
    return;
  }
  if (event.key !== "Enter" || event.shiftKey) return;
  if (target.tagName === "TEXTAREA") event.preventDefault();
  if (scope === "new") {
    event.preventDefault();
    saveNewLostFound().catch((error) => showToast(`Save failed: ${error.message}`, "error"));
    return;
  }
  if (scope === "edit" && id && state.lostFoundEditingId === id) {
    event.preventDefault();
    saveLostFoundEdit(id).catch((error) => showToast(`Update failed: ${error.message}`, "error"));
  }
}

async function onLostFoundStatusToggle(event) {
  const input = event.target;
  if (!input.matches('input[data-lost-found-action="toggle-status"]')) return;
  const id = clean(input.dataset.id);
  const record = state.lostFound.find((item) => item.id === id);
  if (!record) return;
  const nextStatus = input.checked ? "Closed" : "Open";
  input.disabled = true;
  try {
    await api(`/api/lost-found?id=${encodeURIComponent(id)}`, {
      method: "PUT",
      body: { status: nextStatus },
    });
    await loadLostFound({ silent: true });
    renderLostFound();
    showToast(`Status changed to ${nextStatus}.`, "success");
  } catch (error) {
    input.checked = !input.checked;
    showToast(`Status update failed: ${error.message}`, "error");
  } finally {
    input.disabled = false;
  }
}

async function saveNewLostFound() {
  const draft = state.lostFoundDraft || emptyLostFoundDraft();
  if (!clean(draft.objectDescription)) return showToast("Please fill Object Description.", "error");
  try {
    await api("/api/lost-found", {
      method: "POST",
      body: {
        who_found: clean(draft.whoFound),
        who_recorded: clean(draft.whoRecorded),
        location_found: clean(draft.location),
        object_description: clean(draft.objectDescription),
        notes: clean(draft.notes),
        stored_location: normalizeLostFoundStored(draft.stored),
        status: normalizeStatusUi(draft.status),
      },
    });
    state.lostFoundDraft = emptyLostFoundDraft();
    await loadLostFound();
    renderLostFound();
    showToast("Lost&Found record added.", "success");
  } catch (error) {
    showToast(`Save failed: ${error.message}`, "error");
  }
}

async function saveLostFoundEdit(id) {
  const draft = state.lostFoundEditDraft;
  if (!id || !draft) return;
  if (!clean(draft.objectDescription)) return showToast("Please fill Object Description.", "error");
  try {
    await api(`/api/lost-found?id=${encodeURIComponent(id)}`, {
      method: "PUT",
      body: {
        who_found: clean(draft.whoFound),
        who_recorded: clean(draft.whoRecorded),
        location_found: clean(draft.location),
        object_description: clean(draft.objectDescription),
        notes: clean(draft.notes),
        stored_location: normalizeLostFoundStored(draft.stored),
        status: normalizeStatusUi(draft.status),
      },
    });
    state.lostFoundEditingId = null;
    state.lostFoundEditDraft = null;
    await loadLostFound();
    renderLostFound();
    showToast("Lost&Found record updated.", "success");
  } catch (error) {
    showToast(`Update failed: ${error.message}`, "error");
  }
}

function getFilteredLostFoundRecords() {
  const number = clean(els.lostFoundFilterNumber.value).toLowerCase();
  const date = clean(els.lostFoundFilterDate.value);
  const whoFound = clean(els.lostFoundFilterWhoFound.value).toLowerCase();
  const whoRecorded = clean(els.lostFoundFilterWhoRecorded.value).toLowerCase();
  const where = clean(els.lostFoundFilterWhere.value).toLowerCase();
  const objectDescription = clean(els.lostFoundFilterObject.value).toLowerCase();
  const notes = clean(els.lostFoundFilterNotes.value).toLowerCase();
  const stored = clean(els.lostFoundFilterStored.value);
  const onlyOpen = !!els.lostFoundOnlyOpen.checked;
  return state.lostFound
    .filter((record) => {
      const createdDate = lostFoundTimestampDate(record);
      return (!onlyOpen || !isClosedStatus(record.status)) &&
        (!number || String(record.number).toLowerCase().includes(number)) &&
        (!date || createdDate === date) &&
        (!whoFound || record.whoFound.toLowerCase().includes(whoFound)) &&
        (!whoRecorded || record.whoRecorded.toLowerCase().includes(whoRecorded)) &&
        (!where || record.location.toLowerCase().includes(where)) &&
        (!objectDescription || record.objectDescription.toLowerCase().includes(objectDescription)) &&
        (!notes || record.notes.toLowerCase().includes(notes)) &&
        (!stored || record.stored === stored);
    })
    .sort((a, b) => {
      const at = new Date(clean(a.createdAt)).getTime() || 0;
      const bt = new Date(clean(b.createdAt)).getTime() || 0;
      if (at !== bt) return bt - at;
      return (Number(b.number) || 0) - (Number(a.number) || 0);
    });
}

function getFilteredEntries() {
  const search = clean(els.search.value).toLowerCase();
  const showActive = !!els.showActive.checked;
  const status = clean(els.statusFilter.value);
  const category = clean(els.categoryFilter.value);
  const from = clean(els.fromDate.value);
  const to = clean(els.toDate.value);
  const filtered = state.entries.filter((e) => {
    const text = `${e.person} ${e.message}`.toLowerCase();
    return (!showActive || isEntryActive(e)) &&
      (!search || text.includes(search)) &&
      (!status || e.status === status) &&
      (!category || e.category === category) &&
      (!from || e.date >= from) &&
      (!to || e.date <= to);
  });
  return sortEntries(filtered);
}

function isCommunicationsGroupingEnabled() {
  return !!els.groupCommunications?.checked;
}

function splitCommunicationGroups(rows) {
  const open = [];
  const closed = [];
  rows.forEach((entry) => {
    if (isClosedStatus(entry.status)) closed.push(entry);
    else open.push(entry);
  });
  return [
    { key: "open", label: "Open", rows: open },
    { key: "closed", label: "Closed", rows: closed },
  ].filter((group) => group.rows.length);
}

function communicationGroupLabel(group) {
  const count = Array.isArray(group?.rows) ? group.rows.length : 0;
  return `${clean(group?.label) || "Group"}: ${count} record${count === 1 ? "" : "s"}`;
}

function isEntryActive(entry) {
  if (!isClosedStatus(entry.status)) return true;
  const updated = entryUpdatedTime(entry);
  if (!updated) return false;
  const hours24 = 24 * 60 * 60 * 1000;
  return Date.now() - updated.getTime() <= hours24;
}

function entryCreatedTime(entry) {
  const createdAt = clean(entry?.createdAt);
  if (createdAt) {
    const dt = new Date(createdAt);
    if (!Number.isNaN(dt.getTime())) return dt;
  }
  const date = clean(entry?.date);
  const time = normalizeTime(clean(entry?.time)) || "00:00";
  if (!date) return null;
  const fallback = new Date(`${date}T${time}:00`);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function entryUpdatedTime(entry) {
  const updatedAt = clean(entry?.updatedAt);
  if (updatedAt) {
    const dt = new Date(updatedAt);
    if (!Number.isNaN(dt.getTime())) return dt;
  }
  return entryCreatedTime(entry);
}

function entryDateTimeForSort(entry) {
  const date = clean(entry?.date);
  const time = normalizeTime(clean(entry?.time)) || "00:00";
  if (date) {
    const dt = new Date(`${date}T${time}:00`);
    if (!Number.isNaN(dt.getTime())) return dt;
  }
  return entryCreatedTime(entry);
}

function onSortToggle(event) {
  const button = event.target.closest("button[data-sort]");
  if (!button) return;
  const key = clean(button.dataset.sort);
  if (!key) return;
  if (state.sort.key === key) {
    state.sort.dir = state.sort.dir === "asc" ? "desc" : "asc";
  } else {
    state.sort.key = key;
    state.sort.dir = key === "date" ? "desc" : "asc";
  }
  render();
}

function resetSortDefault() {
  state.sort.key = "date";
  state.sort.dir = "desc";
}

function onRowKeydown(event) {
  const target = event.target;
  const scope = clean(target?.dataset?.scope);
  const id = clean(target?.dataset?.id);
  if (!scope) return;

  if (event.key === "Escape" && scope === "edit") {
    event.preventDefault();
    state.editingId = null;
    state.editDraft = null;
    showToast("Edit canceled.", "info");
    render();
    return;
  }

  if (event.key !== "Enter" || event.shiftKey) return;
  if (target.tagName === "TEXTAREA") event.preventDefault();
  if (scope === "new") {
    event.preventDefault();
    saveNew().catch((e) => showToast(`Save failed: ${e.message}`, "error"));
    return;
  }
  if (scope === "edit" && id && state.editingId === id) {
    event.preventDefault();
    saveEdit(id).catch((e) => showToast(`Update failed: ${e.message}`, "error"));
  }
}

function sortEntries(entries) {
  const key = state.sort.key;
  const dir = state.sort.dir === "asc" ? 1 : -1;
  const getValue = (entry) => {
    if (key === "date") {
      const dt = entryDateTimeForSort(entry);
      return dt ? dt.getTime() : 0;
    }
    if (key === "person") return clean(entry.person).toLowerCase();
    if (key === "status") return isClosedStatus(entry.status) ? 1 : 0;
    if (key === "category") return clean(entry.category).toLowerCase();
    return "";
  };
  return [...entries].sort((a, b) => {
    const av = getValue(a);
    const bv = getValue(b);
    if (av < bv) return -1 * dir;
    if (av > bv) return 1 * dir;
    return clean(a.id).localeCompare(clean(b.id)) * dir;
  });
}

function updateSortIndicators() {
  const buttons = els.tableHead.querySelectorAll("button[data-sort]");
  buttons.forEach((button) => {
    const key = clean(button.dataset.sort);
    const active = key === state.sort.key;
    const indicator = button.querySelector(".sort-indicator");
    button.classList.toggle("active", active);
    if (indicator) indicator.textContent = active ? (state.sort.dir === "asc" ? "↑" : "↓") : "";
  });
}

function syncStickyRows() {
  const row = els.tableHead.querySelector("tr");
  if (!row) return;
  const height = Math.max(36, row.getBoundingClientRect().height);
  els.tableWrap.style.setProperty("--table-head-height", `${height}px`);
}

function emptyServiceDraft() {
  return {
    id: "",
    requestNumber: "",
    serviceType: "",
    customerName: "",
    customerEmail: "",
    customerPhone: "",
    pax: 1,
    notes: "",
    date: "",
    time: "",
    pickupLocation: "",
    dropoffLocation: "",
    flightNumber: "",
    hasReturn: false,
    returnPickup: "",
    returnDropoff: "",
    returnDate: "",
    returnTime: "",
    returnFlight: "",
    price: 0,
    priceManual: false,
    status: "Submitted",
    providerUserId: "",
    providerEmail: "",
    language: "en",
    audit: [],
    createdAt: "",
    updatedAt: "",
  };
}

function normalizeServiceStatus(value) {
  const raw = clean(value).toLowerCase();
  if (raw === "approved") return "Approved";
  if (raw === "cancelled" || raw === "canceled") return "Cancelled";
  if (raw === "completed") return "Completed";
  return "Submitted";
}

function normalizeServiceBool(value) {
  if (typeof value === "boolean") return value;
  const raw = clean(value).toLowerCase();
  return ["true", "1", "yes"].includes(raw);
}

function normalizeServicePriceMode(value) {
  return clean(value) === "airport_matrix" ? "airport_matrix" : "open";
}

function normalizeServiceAudit(audit) {
  return (Array.isArray(audit) ? audit : [])
    .map((item) => ({
      at: clean(item?.at),
      action: clean(item?.action),
      user: clean(item?.user),
      summary: clean(item?.summary),
    }))
    .filter((item) => item.at && item.action)
    .slice(-50);
}

function draftText(value) {
  return String(value ?? "");
}

function sanitizeServiceConfigClient(item = {}) {
  const serviceType = clean(item.serviceType || item.service_type);
  const airportTransfer = normalizeServiceBool(item.airportTransfer ?? item.airport_transfer);
  return {
    id: clean(item.id) || serviceType.toLowerCase().replace(/[^a-z0-9]+/g, "-"),
    serviceType,
    providerUserId: clean(item.providerUserId || item.provider_user_id),
    providerEmail: clean(item.providerEmail || item.provider_email).toLowerCase(),
    airportTransfer,
    hasReturn: normalizeServiceBool(item.hasReturn ?? item.has_return),
    approvedByDefault: normalizeServiceBool(item.approvedByDefault ?? item.approved_by_default),
    priceMode: normalizeServicePriceMode(item.priceMode || item.price_mode),
    priceMatrix: {
      oneWay: {
        "1-3": Number(normalizeNumber(item?.priceMatrix?.oneWay?.["1-3"] ?? item?.price_matrix?.oneWay?.["1-3"]) || 0),
        "4-7": Number(normalizeNumber(item?.priceMatrix?.oneWay?.["4-7"] ?? item?.price_matrix?.oneWay?.["4-7"]) || 0),
        "8-11": Number(normalizeNumber(item?.priceMatrix?.oneWay?.["8-11"] ?? item?.price_matrix?.oneWay?.["8-11"]) || 0),
        "12-16": Number(normalizeNumber(item?.priceMatrix?.oneWay?.["12-16"] ?? item?.price_matrix?.oneWay?.["12-16"]) || 0),
      },
      returnTrip: {
        "1-3": Number(normalizeNumber(item?.priceMatrix?.returnTrip?.["1-3"] ?? item?.price_matrix?.returnTrip?.["1-3"]) || 0),
        "4-7": Number(normalizeNumber(item?.priceMatrix?.returnTrip?.["4-7"] ?? item?.price_matrix?.returnTrip?.["4-7"]) || 0),
        "8-11": Number(normalizeNumber(item?.priceMatrix?.returnTrip?.["8-11"] ?? item?.price_matrix?.returnTrip?.["8-11"]) || 0),
        "12-16": Number(normalizeNumber(item?.priceMatrix?.returnTrip?.["12-16"] ?? item?.price_matrix?.returnTrip?.["12-16"]) || 0),
      },
    },
    confirmationTemplate: clean(item.confirmationTemplate || item.confirmation_template) || defaultServiceConfirmationTemplate(serviceType || "Service", airportTransfer),
  };
}

function sanitizeServiceSettingsClient(settings) {
  const output = clone(DEFAULT_SERVICE_SETTINGS);
  output.automaticEmailRecipients = parseEmailList(settings?.automaticEmailRecipients || settings?.automatic_email_recipients);
  const liveFlightStatusEnabled = settings?.liveFlightStatusEnabled ?? settings?.live_flight_status_enabled;
  output.liveFlightStatusEnabled = typeof liveFlightStatusEnabled === "boolean"
    ? liveFlightStatusEnabled
    : !["false", "0", "no"].includes(clean(liveFlightStatusEnabled).toLowerCase());
  const configs = Array.isArray(settings?.serviceConfigs || settings?.service_configs)
    ? settings.serviceConfigs || settings.service_configs
    : [];
  output.serviceConfigs = configs.map(sanitizeServiceConfigClient).filter((item) => item.serviceType);
  if (!output.serviceConfigs.length) output.serviceConfigs = clone(DEFAULT_SERVICE_SETTINGS.serviceConfigs);
  return output;
}

function serviceLiveFlightStatusEnabled() {
  if (clean(state.access?.profile?.name).toLowerCase() === "service provider") return false;
  return state.serviceSettings?.liveFlightStatusEnabled !== false;
}

function normalizeServiceConfirmationLanguage(value) {
  return normalizeProposalLanguage(value);
}

function serviceConfigs() {
  return Array.isArray(state.serviceSettings?.serviceConfigs) ? state.serviceSettings.serviceConfigs : [];
}

function serviceConfigByType(serviceType) {
  return serviceConfigs().find((item) => clean(item.serviceType) === clean(serviceType)) || serviceConfigs()[0] || null;
}

const PHONE_COUNTRY_CODE_MAP = [
  ["351", "PT", "Portugal"],
  ["353", "IE", "Ireland"],
  ["354", "IS", "Iceland"],
  ["355", "AL", "Albania"],
  ["356", "MT", "Malta"],
  ["357", "CY", "Cyprus"],
  ["358", "FI", "Finland"],
  ["359", "BG", "Bulgaria"],
  ["370", "LT", "Lithuania"],
  ["371", "LV", "Latvia"],
  ["372", "EE", "Estonia"],
  ["373", "MD", "Moldova"],
  ["374", "AM", "Armenia"],
  ["375", "BY", "Belarus"],
  ["376", "AD", "Andorra"],
  ["377", "MC", "Monaco"],
  ["378", "SM", "San Marino"],
  ["380", "UA", "Ukraine"],
  ["381", "RS", "Serbia"],
  ["382", "ME", "Montenegro"],
  ["383", "XK", "Kosovo"],
  ["385", "HR", "Croatia"],
  ["386", "SI", "Slovenia"],
  ["387", "BA", "Bosnia and Herzegovina"],
  ["389", "MK", "North Macedonia"],
  ["420", "CZ", "Czechia"],
  ["421", "SK", "Slovakia"],
  ["423", "LI", "Liechtenstein"],
  ["43", "AT", "Austria"],
  ["44", "GB", "United Kingdom"],
  ["45", "DK", "Denmark"],
  ["46", "SE", "Sweden"],
  ["47", "NO", "Norway"],
  ["48", "PL", "Poland"],
  ["49", "DE", "Germany"],
  ["30", "GR", "Greece"],
  ["31", "NL", "Netherlands"],
  ["32", "BE", "Belgium"],
  ["33", "FR", "France"],
  ["34", "ES", "Spain"],
  ["36", "HU", "Hungary"],
  ["39", "IT", "Italy"],
  ["41", "CH", "Switzerland"],
  ["40", "RO", "Romania"],
  ["52", "MX", "Mexico"],
  ["53", "CU", "Cuba"],
  ["54", "AR", "Argentina"],
  ["55", "BR", "Brazil"],
  ["56", "CL", "Chile"],
  ["57", "CO", "Colombia"],
  ["58", "VE", "Venezuela"],
  ["1", "US", "United States / Canada"],
  ["60", "MY", "Malaysia"],
  ["61", "AU", "Australia"],
  ["62", "ID", "Indonesia"],
  ["63", "PH", "Philippines"],
  ["64", "NZ", "New Zealand"],
  ["65", "SG", "Singapore"],
  ["66", "TH", "Thailand"],
  ["81", "JP", "Japan"],
  ["82", "KR", "South Korea"],
  ["84", "VN", "Vietnam"],
  ["86", "CN", "China"],
  ["90", "TR", "Turkey"],
  ["91", "IN", "India"],
  ["92", "PK", "Pakistan"],
  ["93", "AF", "Afghanistan"],
  ["94", "LK", "Sri Lanka"],
  ["95", "MM", "Myanmar"],
  ["98", "IR", "Iran"],
  ["212", "MA", "Morocco"],
  ["213", "DZ", "Algeria"],
  ["216", "TN", "Tunisia"],
  ["218", "LY", "Libya"],
  ["220", "GM", "Gambia"],
  ["221", "SN", "Senegal"],
  ["223", "ML", "Mali"],
  ["225", "CI", "Cote d'Ivoire"],
  ["226", "BF", "Burkina Faso"],
  ["227", "NE", "Niger"],
  ["228", "TG", "Togo"],
  ["229", "BJ", "Benin"],
  ["230", "MU", "Mauritius"],
  ["231", "LR", "Liberia"],
  ["232", "SL", "Sierra Leone"],
  ["233", "GH", "Ghana"],
  ["234", "NG", "Nigeria"],
  ["238", "CV", "Cape Verde"],
  ["239", "ST", "Sao Tome and Principe"],
  ["240", "GQ", "Equatorial Guinea"],
  ["241", "GA", "Gabon"],
  ["242", "CG", "Republic of the Congo"],
  ["243", "CD", "Democratic Republic of the Congo"],
  ["244", "AO", "Angola"],
  ["250", "RW", "Rwanda"],
  ["251", "ET", "Ethiopia"],
  ["254", "KE", "Kenya"],
  ["255", "TZ", "Tanzania"],
  ["256", "UG", "Uganda"],
  ["260", "ZM", "Zambia"],
  ["261", "MG", "Madagascar"],
  ["263", "ZW", "Zimbabwe"],
  ["264", "NA", "Namibia"],
  ["265", "MW", "Malawi"],
  ["266", "LS", "Lesotho"],
  ["267", "BW", "Botswana"],
  ["268", "SZ", "Eswatini"],
  ["269", "KM", "Comoros"],
  ["27", "ZA", "South Africa"],
  ["971", "AE", "United Arab Emirates"],
  ["972", "IL", "Israel"],
  ["973", "BH", "Bahrain"],
  ["974", "QA", "Qatar"],
  ["965", "KW", "Kuwait"],
  ["966", "SA", "Saudi Arabia"],
  ["20", "EG", "Egypt"],
].sort((a, b) => b[0].length - a[0].length);

function isValidInternationalPhone(value) {
  const raw = clean(value);
  if (!raw) return true;
  if (!/^\+[0-9][0-9\s().-]{5,20}$/.test(raw)) return false;
  const digits = raw.replace(/\D/g, "");
  return digits.length >= 6 && digits.length <= 15;
}

function normalizeFlightCode(value) {
  return clean(value).toUpperCase().replace(/\s+/g, "");
}

function formatPredictionTime(value, timeZone = "") {
  const raw = clean(value);
  if (!raw) return "";
  if (/^\d{2}:\d{2}$/.test(raw)) return raw;
  const isoClockMatch = raw.match(/T(\d{2}:\d{2})(?::\d{2})?(?:[.,]\d+)?(?:Z|[+-]\d{2}:\d{2})?$/);
  if (isoClockMatch) return isoClockMatch[1];
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return "";
  try {
    return new Intl.DateTimeFormat("pt-PT", {
      timeZone: "Europe/Lisbon",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    }).format(date);
  } catch {
    return raw.slice(11, 16);
  }
}

function formatFlightStatusLabel(value) {
  const raw = clean(value).replace(/[_-]+/g, " ").trim();
  if (!raw) return "";
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}

function serviceDraftPredictionKey({ flightNumber = "", date = "", leg = "arrival" }) {
  const flight = normalizeFlightCode(flightNumber);
  const when = clean(date);
  if (!flight || !when) return "";
  return `${when}|${flight}|${leg === "departure" ? "departure" : "arrival"}`;
}

function serviceAirportPattern() {
  return /(airport|aeroporto|humberto delgado|portela|lisbon airport|aeroporto de lisboa|terminal 1|terminal 2)/i;
}

function serviceFlightLegForLocations(pickupLocation, dropoffLocation) {
  const airportPattern = serviceAirportPattern();
  const pickup = clean(pickupLocation);
  const dropoff = clean(dropoffLocation);
  if (airportPattern.test(pickup)) return "arrival";
  if (airportPattern.test(dropoff)) return "departure";
  return "";
}

function setServiceDraftPrediction(slot, key, text) {
  if (!state.serviceDraftFlightPredictions?.[slot]) return;
  state.serviceDraftFlightPredictions[slot] = { key, text };
}

function resetServiceDraftPredictionState({ keepCache = true } = {}) {
  if (state.serviceDraftFlightPredictions.timer) {
    clearTimeout(state.serviceDraftFlightPredictions.timer);
    state.serviceDraftFlightPredictions.timer = null;
  }
  if (!keepCache) state.serviceDraftFlightPredictions.cache = {};
  state.serviceDraftFlightPredictions.main = { key: "", text: "" };
  state.serviceDraftFlightPredictions.return = { key: "", text: "" };
  renderServiceDraftPredictionHints();
}

function renderServiceDraftPredictionHints() {
  const mainText = clean(state.serviceDraftFlightPredictions.main?.text);
  if (els.serviceTimePrediction) {
    els.serviceTimePrediction.hidden = !mainText;
    els.serviceTimePrediction.textContent = mainText;
  }
  const returnText = clean(state.serviceDraftFlightPredictions.return?.text);
  if (els.serviceReturnTimePrediction) {
    els.serviceReturnTimePrediction.hidden = !returnText;
    els.serviceReturnTimePrediction.textContent = returnText;
  }
}

async function fetchServiceDraftPrediction(slot, options) {
  const key = serviceDraftPredictionKey(options);
  if (!key) {
    setServiceDraftPrediction(slot, "", "");
    renderServiceDraftPredictionHints();
    return;
  }
  const cached = state.serviceDraftFlightPredictions.cache[key];
  if (cached?.loaded) {
    setServiceDraftPrediction(slot, key, cached.text);
    renderServiceDraftPredictionHints();
    return;
  }
  if (cached?.pending) {
    setServiceDraftPrediction(slot, key, "checking...");
    renderServiceDraftPredictionHints();
    return;
  }
  state.serviceDraftFlightPredictions.cache[key] = { text: "checking...", loaded: false, pending: true };
  setServiceDraftPrediction(slot, key, "checking...");
  renderServiceDraftPredictionHints();
  try {
    const result = await api(`/api/aviationstack-flight?flight=${encodeURIComponent(normalizeFlightCode(options.flightNumber))}&date=${encodeURIComponent(clean(options.date))}&leg=${encodeURIComponent(options.leg === "departure" ? "departure" : "arrival")}&time_kind=scheduled&pickup=${encodeURIComponent(clean(options.pickupLocation))}&dropoff=${encodeURIComponent(clean(options.dropoffLocation))}`);
    const timeText = formatPredictionTime(result?.predictedTime, result?.timeZone);
    const label = options.leg === "departure" ? "ETD" : "ETA";
    const statusLabel = formatFlightStatusLabel(result?.status);
    const displayText = `${statusLabel ? `${statusLabel} ` : ""}${timeText ? `${label} ${timeText}` : `${label} -`}`.trim();
    state.serviceDraftFlightPredictions.cache[key] = { text: displayText, loaded: true, pending: false };
  } catch (error) {
    const message = clean(error.message);
    let text = "lookup unavailable";
    if (/Missing server environment variable: AVIATIONSTACK_API_KEY/i.test(message)) text = "Aviationstack not configured";
    else if (/not found/i.test(message)) text = "not found";
    state.serviceDraftFlightPredictions.cache[key] = { text, loaded: true, pending: false };
  }
  if (state.serviceDraftFlightPredictions?.[slot]?.key === key || !clean(state.serviceDraftFlightPredictions?.[slot]?.key)) {
    setServiceDraftPrediction(slot, key, state.serviceDraftFlightPredictions.cache[key].text);
    renderServiceDraftPredictionHints();
  }
}

function queueServiceDraftPredictionRefresh() {
  if (state.serviceDraftFlightPredictions.timer) {
    clearTimeout(state.serviceDraftFlightPredictions.timer);
    state.serviceDraftFlightPredictions.timer = null;
  }
  const draft = state.serviceDraft;
  const config = serviceConfigByType(draft.serviceType);
  const showFlight = serviceUsesFlightFields(config, draft.serviceType);
  if (!showFlight) {
    setServiceDraftPrediction("main", "", "");
    setServiceDraftPrediction("return", "", "");
    renderServiceDraftPredictionHints();
    return;
  }
  const mainOptions = {
    flightNumber: draft.flightNumber,
    date: draft.date,
    leg: serviceFlightLegForLocations(draft.pickupLocation, draft.dropoffLocation) || "arrival",
    pickupLocation: draft.pickupLocation,
    dropoffLocation: draft.dropoffLocation,
  };
  const returnOptions = {
    flightNumber: draft.returnFlight,
    date: draft.returnDate,
    leg: serviceFlightLegForLocations(draft.returnPickup, draft.returnDropoff) || "departure",
    pickupLocation: draft.returnPickup,
    dropoffLocation: draft.returnDropoff,
  };
  setServiceDraftPrediction("main", serviceDraftPredictionKey(mainOptions), clean(draft.flightNumber) && clean(draft.date) ? "checking..." : "");
  setServiceDraftPrediction("return", serviceDraftPredictionKey(returnOptions), draft.hasReturn && clean(draft.returnFlight) && clean(draft.returnDate) ? "checking..." : "");
  renderServiceDraftPredictionHints();
  state.serviceDraftFlightPredictions.timer = window.setTimeout(() => {
    fetchServiceDraftPrediction("main", mainOptions);
    if (draft.hasReturn) fetchServiceDraftPrediction("return", returnOptions);
    else {
      setServiceDraftPrediction("return", "", "");
      renderServiceDraftPredictionHints();
    }
  }, 350);
}

function formatServiceDateInput(value) {
  const raw = clean(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  return raw;
}

function parseServiceDateInput(value) {
  const raw = clean(value);
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const match = raw.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (!match) return "";
  const year = match[1];
  const month = match[2].padStart(2, "0");
  const day = match[3].padStart(2, "0");
  const iso = `${year}-${month}-${day}`;
  const dt = new Date(iso);
  if (Number.isNaN(dt.getTime())) return "";
  return formatDate(dt) === iso ? iso : "";
}

function syncServiceDatePickers() {
  if (els.serviceDatePicker) els.serviceDatePicker.value = state.serviceDraft.date || "";
  if (els.serviceReturnDatePicker) {
    els.serviceReturnDatePicker.value = state.serviceDraft.returnDate || "";
    els.serviceReturnDatePicker.min = state.serviceDraft.date || "";
  }
}

function servicePhoneCountryInfo(value) {
  const digits = clean(value).replace(/\D/g, "");
  if (!digits) return { isoCode: "", title: "Add country code to show flag" };
  const match = PHONE_COUNTRY_CODE_MAP.find(([code]) => digits.startsWith(code));
  if (!match) return { isoCode: "", title: `Unknown country code (+${digits.slice(0, Math.min(4, digits.length))})` };
  const [code, isoCode, label] = match;
  return {
    isoCode: clean(isoCode).toLowerCase(),
    title: `${label} (+${code})`,
  };
}

function renderServicePhoneFlag() {
  if (!els.serviceCustomerPhoneFlag) return;
  const info = servicePhoneCountryInfo(els.serviceCustomerPhone?.value || state.serviceDraft?.customerPhone);
  els.serviceCustomerPhoneFlag.title = info.title;
  els.serviceCustomerPhoneFlag.setAttribute("aria-label", info.title);
  if (!info.isoCode || info.isoCode === "xk") {
    els.serviceCustomerPhoneFlag.innerHTML = '<span class="phone-flag-fallback">🌐</span>';
    return;
  }
  els.serviceCustomerPhoneFlag.innerHTML = `<img src="https://flagcdn.com/24x18/${info.isoCode}.png" alt="" width="24" height="18" loading="lazy" referrerpolicy="no-referrer" onerror="this.parentElement.innerHTML='&lt;span class=&quot;phone-flag-fallback&quot;&gt;🌐&lt;/span&gt;';" />`;
}

function serviceAirportArrivalCandidate(row) {
  const today = formatDate(new Date());
  if (!clean(row.flightNumber) || clean(row.date) !== today) return false;
  return !!serviceFlightLegForLocations(row.pickupLocation, row.dropoffLocation);
}

function serviceFlightStatusKey(row) {
  return `${clean(row.date)}|${normalizeFlightCode(row.flightNumber)}|${clean(row.pickupLocation)}|${clean(row.dropoffLocation)}`;
}

function serviceFlightStatusText(row) {
  const cached = state.serviceFlightStatuses.cache[serviceFlightStatusKey(row)];
  return clean(cached?.text);
}

function renderServiceFlightCell(row) {
  const flight = clean(row.flightNumber) || "-";
  if (!serviceLiveFlightStatusEnabled()) return escape(flight);
  const statusText = serviceFlightStatusText(row);
  if (!statusText) return escape(flight);
  return `${escape(flight)}<small class="service-flight-status-note">(${escape(statusText)})</small>`;
}

async function fetchServiceFlightStatus(row) {
  if (!serviceLiveFlightStatusEnabled()) return;
  const key = serviceFlightStatusKey(row);
  if (!serviceAirportArrivalCandidate(row)) {
    delete state.serviceFlightStatuses.cache[key];
    return;
  }
  if (state.serviceFlightStatuses.cache[key]?.loaded || state.serviceFlightStatuses.cache[key]?.pending) return;
  state.serviceFlightStatuses.cache[key] = { text: "checking...", loaded: false, pending: true };
  renderServices();
  try {
    const leg = serviceFlightLegForLocations(row.pickupLocation, row.dropoffLocation) || "arrival";
    const result = await api(`/api/aviationstack-flight?flight=${encodeURIComponent(normalizeFlightCode(row.flightNumber))}&date=${encodeURIComponent(clean(row.date))}&leg=${encodeURIComponent(leg)}&pickup=${encodeURIComponent(clean(row.pickupLocation))}&dropoff=${encodeURIComponent(clean(row.dropoffLocation))}`);
    const timeText = formatPredictionTime(result?.predictedTime, result?.timeZone);
    const label = leg === "departure" ? "ETD" : "ETA";
    const statusLabel = formatFlightStatusLabel(result?.status);
    const displayText = `${statusLabel ? `${statusLabel} ` : ""}${timeText ? `${label} ${timeText}` : `${label} -`}`.trim();
    state.serviceFlightStatuses.cache[key] = { text: displayText, loaded: true, pending: false };
  } catch (error) {
    const message = clean(error.message);
    if (/Missing server environment variable: AVIATIONSTACK_API_KEY/i.test(message)) {
      state.serviceFlightStatuses.cache[key] = { text: "Aviationstack not configured", loaded: true, pending: false };
    } else if (/not found/i.test(message)) {
      state.serviceFlightStatuses.cache[key] = { text: "not found", loaded: true, pending: false };
    } else {
      state.serviceFlightStatuses.cache[key] = { text: "lookup unavailable", loaded: true, pending: false };
    }
  }
  renderServices();
}

function refreshVisibleServiceStatuses() {
  if (!serviceLiveFlightStatusEnabled()) {
    state.serviceFlightStatuses.cache = {};
    if (state.serviceFlightStatuses.timer) {
      clearTimeout(state.serviceFlightStatuses.timer);
      state.serviceFlightStatuses.timer = null;
    }
    state.serviceFlightStatuses.initialized = false;
    return;
  }
  if (state.serviceFlightStatuses.initialized) return;
  state.serviceFlightStatuses.initialized = true;
  if (state.serviceFlightStatuses.timer) {
    clearTimeout(state.serviceFlightStatuses.timer);
    state.serviceFlightStatuses.timer = null;
  }
  state.serviceFlightStatuses.timer = window.setTimeout(() => {
    const rows = getFilteredServices().filter(serviceAirportArrivalCandidate);
    rows.forEach((row) => {
      fetchServiceFlightStatus(row);
    });
  }, 250);
}

function ensureServiceSettingsTemplateType() {
  const configs = serviceConfigs();
  if (!configs.length) {
    state.serviceSettingsTemplateType = "";
    return "";
  }
  const current = clean(state.serviceSettingsTemplateType);
  if (configs.some((item) => clean(item.id) === current)) return current;
  state.serviceSettingsTemplateType = clean(configs[0].id);
  return state.serviceSettingsTemplateType;
}

function currentServiceSettingsTemplateConfig() {
  const selectedId = ensureServiceSettingsTemplateType();
  return serviceConfigs().find((item) => clean(item.id) === selectedId) || serviceConfigs()[0] || null;
}

function serviceConfirmationTemplate(config, language = "en") {
  const normalizedLanguage = normalizeServiceConfirmationLanguage(language);
  if (!config) return defaultServiceConfirmationTemplate("Service", false);
  if ((normalizedLanguage === "pt" || normalizedLanguage === "es") && SERVICE_CONFIRMATION_TEMPLATES[normalizedLanguage]) {
    return SERVICE_CONFIRMATION_TEMPLATES[normalizedLanguage](!!config.airportTransfer);
  }
  return clean(config.confirmationTemplate) || defaultServiceConfirmationTemplate(config.serviceType || "Service", !!config.airportTransfer);
}

function serviceUsesFlightFields(config, serviceType) {
  return !!config?.airportTransfer;
}

function serviceBandForPax(pax) {
  const count = Math.max(1, Math.round(Number(pax || 0)));
  if (count <= 3) return "1-3";
  if (count <= 7) return "4-7";
  if (count <= 11) return "8-11";
  return "12-16";
}

function serviceComputedPrice(config, pax, hasReturn) {
  if (!config || normalizeServicePriceMode(config.priceMode) !== "airport_matrix") return null;
  const band = serviceBandForPax(pax);
  return Number(config?.priceMatrix?.[hasReturn ? "returnTrip" : "oneWay"]?.[band] || 0);
}

function serviceProviderOptions(selectedValue = "") {
  const current = clean(selectedValue);
  const options = ['<option value="">Select provider</option>'];
  state.serviceProviders.forEach((provider) => {
    const value = clean(provider.id);
    options.push(`<option value="${escape(value)}" ${value === current ? "selected" : ""}>${escape(provider.email)}</option>`);
  });
  return options.join("");
}

function serviceConfirmationTableRows(draft, config = serviceConfigByType(draft?.serviceType)) {
  const rows = [];
  const push = (label, value, include = true) => {
    const text = clean(value);
    if (!include) return;
    rows.push([label, text || "-"]);
  };
  push("Service", clean(draft.serviceType));
  push("Request #", clean(draft.requestNumber));
  push("Date", formatGroupDateDisplay(draft.date));
  push("Time", clean(draft.time));
  push("Client name", clean(draft.customerName));
  push("Client e-mail", clean(draft.customerEmail), !!clean(draft.customerEmail));
  push("Client phone", clean(draft.customerPhone), !!clean(draft.customerPhone));
  push("Nº of persons", String(draft.pax || 0));
  push("Price", formatMoney(draft.price));
  push("Pick up location", clean(draft.pickupLocation), !!clean(draft.pickupLocation));
  push("Dropoff location", clean(draft.dropoffLocation), !!clean(draft.dropoffLocation));
  push("Flight number", clean(draft.flightNumber), !!config?.airportTransfer);
  push("Status", clean(draft.status), !!clean(draft.status));
  if (draft.hasReturn) {
    push("Return service", "Yes");
    push("Return Date", formatGroupDateDisplay(draft.returnDate), !!clean(draft.returnDate));
    push("Return Time", clean(draft.returnTime), !!clean(draft.returnTime));
    push("Return pickup", clean(draft.returnPickup), !!clean(draft.returnPickup));
    push("Return dropoff", clean(draft.returnDropoff), !!clean(draft.returnDropoff));
    push("Return flight number", clean(draft.returnFlight), !!config?.airportTransfer);
  }
  return rows;
}

function serviceConfirmationTableText(draft, config = serviceConfigByType(draft?.serviceType)) {
  return serviceConfirmationTableRows(draft, config)
    .map(([label, value]) => `${label}: ${value}`)
    .join("\n");
}

function serviceConfirmationTableHtml(draft, config = serviceConfigByType(draft?.serviceType)) {
  const labelStyle = "border: 1pt solid #3F96AA; background: #D8EEF2; padding: 5px 8px; width: 210px; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; font-weight: bold; color: #000000; vertical-align: top;";
  const valueStyle = "border: 1pt solid #3F96AA; padding: 5px 8px; width: 280px; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; color: #000000; vertical-align: top;";
  const rows = serviceConfirmationTableRows(draft, config)
    .map(([label, value]) => `<tr><td style="${labelStyle}">${escape(label)}</td><td style="${valueStyle}">${escape(value).replace(/\n/g, "<br>")}</td></tr>`)
    .join("");
  return `<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin: 12px 0 18px 47px; width: 490px; table-layout: fixed; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; border: 1pt solid #3F96AA;"><tbody>${rows}</tbody></table>`;
}

function serviceConfirmationReplacements(draft, config = serviceConfigByType(draft?.serviceType)) {
  return {
    customer_name: clean(draft.customerName) || "Guest",
    customer_email: clean(draft.customerEmail) || "-",
    customer_phone: clean(draft.customerPhone) || "-",
    service_type: clean(draft.serviceType) || "-",
    request_number: clean(draft.requestNumber) || "-",
    date: formatGroupDateDisplay(draft.date),
    time: clean(draft.time) || "-",
    pax: String(draft.pax || 0),
    pickup_location: clean(draft.pickupLocation) || "-",
    dropoff_location: clean(draft.dropoffLocation) || "-",
    flight_number: clean(draft.flightNumber) || "-",
    return_date: clean(draft.returnDate) ? formatGroupDateDisplay(draft.returnDate) : "-",
    return_time: clean(draft.returnTime) || "-",
    return_pickup: clean(draft.returnPickup) || "-",
    return_dropoff: clean(draft.returnDropoff) || "-",
    return_flight: clean(draft.returnFlight) || "-",
    price: formatMoney(draft.price),
    status: clean(draft.status) || "-",
    service_provider: clean(draft.providerEmail) || "-",
    service_table: serviceConfirmationTableText(draft, config),
  };
}

function serviceConfirmationEmailText(draft, template = serviceConfirmationTemplate(serviceConfigByType(draft?.serviceType))) {
  const config = serviceConfigByType(draft?.serviceType);
  const replacements = serviceConfirmationReplacements(draft, config);
  return Object.entries(replacements).reduce(
    (text, [key, value]) => text.replaceAll(`{{${key}}}`, value),
    template
  );
}

function serviceConfirmationEmailHtmlFromTemplate(draft, template, config = serviceConfigByType(draft?.serviceType)) {
  const replacements = serviceConfirmationReplacements(draft, config);
  const textWithValues = Object.entries(replacements)
    .filter(([key]) => key !== "service_table")
    .reduce((text, [key, value]) => text.replaceAll(`{{${key}}}`, value), template);
  const chunks = textWithValues.split("{{service_table}}").map(groupProposalTextChunkHtml);
  return `<div class="proposal-email-document" style="font-family: Calibri, Arial, Helvetica, sans-serif; color: #000000; font-size: 11pt; line-height: 1.15; max-width: 700px;">${chunks.join(serviceConfirmationTableHtml(draft, config))}</div>`;
}

function serviceConfirmationEmailHtml(draft) {
  const config = serviceConfigByType(draft?.serviceType);
  return serviceConfirmationEmailHtmlFromTemplate(draft, serviceConfirmationTemplate(config), config);
}

function serviceConfirmationLabels(language = "en") {
  const current = normalizeServiceConfirmationLanguage(language);
  if (current === "pt") {
    return {
      service: "Servico",
      requestNumber: "Pedido #",
      date: "Data",
      time: "Hora",
      customerName: "Nome do cliente",
      customerEmail: "E-mail do cliente",
      customerPhone: "Telefone do cliente",
      pax: "Nr de pessoas",
      price: "Preco",
      pickupLocation: "Local de recolha",
      dropoffLocation: "Destino",
      flightNumber: "Nr de voo",
      status: "Estado",
      returnService: "Servico de regresso",
      yes: "Sim",
      returnDate: "Data de regresso",
      returnTime: "Hora de regresso",
      returnPickup: "Recolha regresso",
      returnDropoff: "Destino regresso",
      returnFlightNumber: "Nr voo regresso",
    };
  }
  if (current === "es") {
    return {
      service: "Servicio",
      requestNumber: "Solicitud #",
      date: "Fecha",
      time: "Hora",
      customerName: "Nombre del cliente",
      customerEmail: "E-mail del cliente",
      customerPhone: "Telefono del cliente",
      pax: "Nr de personas",
      price: "Precio",
      pickupLocation: "Lugar de recogida",
      dropoffLocation: "Destino",
      flightNumber: "Nr de vuelo",
      status: "Estado",
      returnService: "Servicio de regreso",
      yes: "Si",
      returnDate: "Fecha de regreso",
      returnTime: "Hora de regreso",
      returnPickup: "Recogida regreso",
      returnDropoff: "Destino regreso",
      returnFlightNumber: "Nr vuelo regreso",
    };
  }
  return {
    service: "Service",
    requestNumber: "Request #",
    date: "Date",
    time: "Time",
    customerName: "Client name",
    customerEmail: "Client e-mail",
    customerPhone: "Client phone",
    pax: "Nr of persons",
    price: "Price",
    pickupLocation: "Pick up location",
    dropoffLocation: "Dropoff location",
    flightNumber: "Flight number",
    status: "Status",
    returnService: "Return service",
    yes: "Yes",
    returnDate: "Return Date",
    returnTime: "Return Time",
    returnPickup: "Return pickup",
    returnDropoff: "Return dropoff",
    returnFlightNumber: "Return flight number",
  };
}

function serviceConfirmationTableRows(draft, config = serviceConfigByType(draft?.serviceType), language = state.serviceConfirmationLanguage) {
  const labels = serviceConfirmationLabels(language);
  const rows = [];
  const push = (label, value, include = true) => {
    const text = clean(value);
    if (!include) return;
    rows.push([label, text || "-"]);
  };
  push(labels.service, clean(draft.serviceType));
  push(labels.requestNumber, clean(draft.requestNumber));
  push(labels.date, formatGroupDateDisplay(draft.date));
  push(labels.time, clean(draft.time));
  push(labels.customerName, clean(draft.customerName));
  push(labels.customerEmail, clean(draft.customerEmail), !!clean(draft.customerEmail));
  push(labels.customerPhone, clean(draft.customerPhone), !!clean(draft.customerPhone));
  push(labels.pax, String(draft.pax || 0));
  push(labels.price, formatMoney(draft.price));
  push(labels.pickupLocation, clean(draft.pickupLocation), !!clean(draft.pickupLocation));
  push(labels.dropoffLocation, clean(draft.dropoffLocation), !!clean(draft.dropoffLocation));
  push(labels.flightNumber, clean(draft.flightNumber), !!config?.airportTransfer);
  push(labels.status, clean(draft.status), !!clean(draft.status));
  if (draft.hasReturn) {
    push(labels.returnService, labels.yes);
    push(labels.returnDate, formatGroupDateDisplay(draft.returnDate), !!clean(draft.returnDate));
    push(labels.returnTime, clean(draft.returnTime), !!clean(draft.returnTime));
    push(labels.returnPickup, clean(draft.returnPickup), !!clean(draft.returnPickup));
    push(labels.returnDropoff, clean(draft.returnDropoff), !!clean(draft.returnDropoff));
    push(labels.returnFlightNumber, clean(draft.returnFlight), !!config?.airportTransfer);
  }
  return rows;
}

function serviceConfirmationTableText(draft, config = serviceConfigByType(draft?.serviceType), language = state.serviceConfirmationLanguage) {
  return serviceConfirmationTableRows(draft, config, language)
    .map(([label, value]) => `${label}: ${value}`)
    .join("\n");
}

function serviceConfirmationTableHtml(draft, config = serviceConfigByType(draft?.serviceType), language = state.serviceConfirmationLanguage) {
  const labelStyle = "border: 1pt solid #3F96AA; background: #D8EEF2; padding: 5px 8px; width: 210px; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; font-weight: bold; color: #000000; vertical-align: top;";
  const valueStyle = "border: 1pt solid #3F96AA; padding: 5px 8px; width: 280px; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; color: #000000; vertical-align: top;";
  const rows = serviceConfirmationTableRows(draft, config, language)
    .map(([label, value]) => `<tr><td style="${labelStyle}">${escape(label)}</td><td style="${valueStyle}">${escape(value).replace(/\n/g, "<br>")}</td></tr>`)
    .join("");
  return `<table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin: 12px 0 18px 47px; width: 490px; table-layout: fixed; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; border: 1pt solid #3F96AA;"><tbody>${rows}</tbody></table>`;
}

function serviceConfirmationReplacements(draft, config = serviceConfigByType(draft?.serviceType), language = state.serviceConfirmationLanguage) {
  return {
    customer_name: clean(draft.customerName) || "Guest",
    customer_email: clean(draft.customerEmail) || "-",
    customer_phone: clean(draft.customerPhone) || "-",
    service_type: clean(draft.serviceType) || "-",
    request_number: clean(draft.requestNumber) || "-",
    date: formatGroupDateDisplay(draft.date),
    time: clean(draft.time) || "-",
    pax: String(draft.pax || 0),
    pickup_location: clean(draft.pickupLocation) || "-",
    dropoff_location: clean(draft.dropoffLocation) || "-",
    flight_number: clean(draft.flightNumber) || "-",
    return_date: clean(draft.returnDate) ? formatGroupDateDisplay(draft.returnDate) : "-",
    return_time: clean(draft.returnTime) || "-",
    return_pickup: clean(draft.returnPickup) || "-",
    return_dropoff: clean(draft.returnDropoff) || "-",
    return_flight: clean(draft.returnFlight) || "-",
    price: formatMoney(draft.price),
    status: clean(draft.status) || "-",
    service_provider: clean(draft.providerEmail) || "-",
    service_table: serviceConfirmationTableText(draft, config, language),
  };
}

function serviceConfirmationEmailText(draft, template = serviceConfirmationTemplate(serviceConfigByType(draft?.serviceType), draft?.language || state.serviceConfirmationLanguage), language = draft?.language || state.serviceConfirmationLanguage) {
  const config = serviceConfigByType(draft?.serviceType);
  const replacements = serviceConfirmationReplacements(draft, config, language);
  return Object.entries(replacements).reduce(
    (text, [key, value]) => text.replaceAll(`{{${key}}}`, value),
    template
  );
}

function serviceConfirmationEmailHtmlFromTemplate(draft, template, config = serviceConfigByType(draft?.serviceType), language = draft?.language || state.serviceConfirmationLanguage) {
  const replacements = serviceConfirmationReplacements(draft, config, language);
  const textWithValues = Object.entries(replacements)
    .filter(([key]) => key !== "service_table")
    .reduce((text, [key, value]) => text.replaceAll(`{{${key}}}`, value), template);
  const chunks = textWithValues.split("{{service_table}}").map(groupProposalTextChunkHtml);
  return `<div class="proposal-email-document" style="font-family: Calibri, Arial, Helvetica, sans-serif; color: #000000; font-size: 11pt; line-height: 1.15; max-width: 700px;">${chunks.join(serviceConfirmationTableHtml(draft, config, language))}</div>`;
}

function serviceConfirmationEmailHtml(draft) {
  const config = serviceConfigByType(draft?.serviceType);
  const language = draft?.language || state.serviceConfirmationLanguage;
  return serviceConfirmationEmailHtmlFromTemplate(draft, serviceConfirmationTemplate(config, language), config, language);
}

function serviceTemplatePreviewDraft(config = currentServiceSettingsTemplateConfig()) {
  const draft = {
    ...emptyServiceDraft(),
    requestNumber: "15",
    serviceType: clean(config?.serviceType) || "Airport Transfer",
    customerName: "Jordyn Piparo",
    customerEmail: "jordynpiparo2015@gmail.com",
    customerPhone: "+15515791054",
    pax: 1,
    date: "2026-04-25",
    time: "23:25",
    pickupLocation: "Aeroporto de Lisboa",
    dropoffLocation: "Lisboa Central Hostel",
    flightNumber: config?.airportTransfer ? "IB0539" : "",
    hasReturn: !!config?.hasReturn,
    returnPickup: "Lisboa Central Hostel",
    returnDropoff: "Aeroporto de Lisboa",
    returnDate: config?.hasReturn ? "2026-05-02" : "",
    returnTime: config?.hasReturn ? "09:30" : "",
    returnFlight: config?.airportTransfer && config?.hasReturn ? "IB0532" : "",
    price: serviceComputedPrice(config, 1, !!config?.hasReturn) ?? 63,
    status: config?.approvedByDefault ? "Approved" : "Submitted",
    providerEmail: clean(config?.providerEmail) || "service@example.com",
    language: normalizeServiceConfirmationLanguage(state.serviceSettingsTemplateLanguage),
  };
  return draft;
}

async function loadServiceSettings({ silent = false } = {}) {
  try {
    const result = await api("/api/service-settings");
    state.serviceSettings = sanitizeServiceSettingsClient(result.settings);
    state.serviceProviders = Array.isArray(result.providers) ? result.providers : [];
    renderServiceSettings();
    if (!silent) setServicesSettingsStatus("Services configuration loaded.");
  } catch (e) {
    state.serviceSettings = clone(DEFAULT_SERVICE_SETTINGS);
    state.serviceProviders = [];
    renderServiceSettings();
    if (!silent) setServicesSettingsStatus(`Using default services configuration (${e.message}).`);
  }
}

async function loadServices({ silent = false, throwOnError = false } = {}) {
  try {
    const result = await api("/api/services");
    state.services = (Array.isArray(result.rows) ? result.rows : []).map(mapServiceRow);
    renderServices();
    if (!silent) setServicesDbStatus(`Loaded ${state.services.length} service request${state.services.length === 1 ? "" : "s"}.`);
  } catch (e) {
    state.services = [];
    renderServices();
    if (!silent) {
      setServicesDbStatus(`Failed to load services: ${e.message}`);
      showToast(`Failed to load services: ${e.message}`, "error");
    }
    if (throwOnError) throw e;
  }
}

function mapServiceRow(row) {
  return {
    id: clean(row.id),
    requestNumber: clean(row.request_number),
    serviceType: clean(row.service_type),
    customerName: clean(row.customer_name),
    customerEmail: clean(row.customer_email),
    customerPhone: clean(row.customer_phone),
    pax: Math.max(1, Math.round(Number(row.pax || 1))),
    notes: clean(row.notes),
    date: clean(row.service_date),
    time: clean(row.service_time),
    pickupLocation: clean(row.pickup_location),
    dropoffLocation: clean(row.dropoff_location),
    flightNumber: clean(row.flight_number),
    hasReturn: !!row.has_return,
    returnPickup: clean(row.return_pickup_location),
    returnDropoff: clean(row.return_dropoff_location),
    returnDate: clean(row.return_date),
    returnTime: clean(row.return_time),
    returnFlight: clean(row.return_flight_number),
    price: Number(normalizeNumber(row.price) || 0),
    status: normalizeServiceStatus(row.status),
    providerUserId: clean(row.provider_user_id),
    providerEmail: clean(row.provider_email),
    language: "en",
    audit: normalizeServiceAudit(row.audit_log),
    createdAt: clean(row.created_at),
    updatedAt: clean(row.updated_at),
  };
}

function expandServiceListRows(service) {
  const base = {
    serviceId: service.id,
    requestNumber: service.requestNumber,
    serviceType: service.serviceType,
    customerName: service.customerName,
    status: service.status,
    pax: service.pax,
    legType: "outbound",
    date: service.date,
    time: service.time,
    pickupLocation: service.pickupLocation,
    flightNumber: service.flightNumber,
    dropoffLocation: service.dropoffLocation,
    price: service.price,
  };
  const rows = [base];
  if (service.hasReturn && clean(service.returnDate)) {
    rows.push({
      serviceId: service.id,
      requestNumber: service.requestNumber,
      serviceType: service.serviceType,
      customerName: service.customerName,
      status: service.status,
      pax: service.pax,
      legType: "return",
      date: service.returnDate,
      time: service.returnTime,
      pickupLocation: service.returnPickup,
      flightNumber: service.returnFlight,
      dropoffLocation: service.returnDropoff,
      price: 0,
    });
  }
  return rows;
}

function serviceStatusTone(status) {
  const normalized = normalizeServiceStatus(status);
  if (normalized === "Approved") return "approved";
  if (normalized === "Cancelled") return "cancelled";
  return "submitted";
}

function canInlineChangeServiceStatus() {
  return clean(state.access?.profile?.name).toLowerCase() === "service provider";
}

function serviceStatusOptions(selected) {
  const normalized = normalizeServiceStatus(selected);
  return ["Submitted", "Approved", "Cancelled"].map((value) => option(value, normalized)).join("");
}

function renderServiceStatusCell(row) {
  const normalized = normalizeServiceStatus(row.status);
  if (!canInlineChangeServiceStatus()) {
    return `<span class="service-status-pill ${serviceStatusTone(normalized)}">${escape(normalized || "-")}</span>`;
  }
  const disabled = state.serviceInlineStatusSaving[row.serviceId] ? "disabled" : "";
  return `<select class="service-inline-status ${serviceStatusTone(normalized)}" data-inline-service-status="${escape(row.serviceId)}" ${disabled}>
    ${serviceStatusOptions(normalized)}
  </select>`;
}

function serviceRelativeDateHint(value) {
  const raw = clean(value);
  if (!raw) return "";
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  const dateOnly = raw.slice(0, 10);
  if (dateOnly === formatDate(today)) return "Today";
  if (dateOnly === formatDate(tomorrow)) return "Tomorrow";
  return "";
}

function renderServiceDateCell(value) {
  const dateText = escape(formatDateOnly(value));
  const hint = serviceRelativeDateHint(value);
  if (!hint) return dateText;
  return `${dateText}<small class="service-date-hint">${escape(hint)}</small>`;
}

function getFilteredServices() {
  const today = formatDate(new Date());
  const filters = state.serviceFilters || {};
  const createdFrom = clean(filters.createdFrom);
  const createdTo = clean(filters.createdTo);
  const serviceDateFrom = clean(filters.dateFrom);
  const serviceDateTo = clean(filters.dateTo);
  const nameNeedle = clean(filters.name).toLowerCase();
  return [...state.services]
    .filter((row) => {
      const createdDate = formatDateInLisbon(row.createdAt);
      if (createdFrom && (!createdDate || createdDate < createdFrom)) return false;
      if (createdTo && (!createdDate || createdDate > createdTo)) return false;
      return true;
    })
    .filter((row) => {
      const mainDate = clean(row.date);
      const returnDate = row.hasReturn ? clean(row.returnDate) : "";
      const matchesMain = (!serviceDateFrom || (mainDate && mainDate >= serviceDateFrom)) && (!serviceDateTo || (mainDate && mainDate <= serviceDateTo));
      const matchesReturn = returnDate && (!serviceDateFrom || returnDate >= serviceDateFrom) && (!serviceDateTo || returnDate <= serviceDateTo);
      return !serviceDateFrom && !serviceDateTo ? true : matchesMain || matchesReturn;
    })
    .filter((row) => !nameNeedle || clean(row.customerName).toLowerCase().includes(nameNeedle))
    .flatMap(expandServiceListRows)
    .filter((row) => !filters.showActive || (["Submitted", "Approved"].includes(normalizeServiceStatus(row.status)) && clean(row.date) >= today))
    .sort((a, b) => {
      const aKey = `${clean(a.date)} ${clean(a.time)} ${clean(a.requestNumber)} ${clean(a.legType)}`;
      const bKey = `${clean(b.date)} ${clean(b.time)} ${clean(b.requestNumber)} ${clean(b.legType)}`;
      return aKey.localeCompare(bKey);
    });
}

function onServiceFilterInput() {
  state.serviceFilters.showActive = !!els.servicesShowActive.checked;
  state.serviceFilters.createdFrom = clean(els.servicesFilterCreatedFrom.value);
  state.serviceFilters.createdTo = clean(els.servicesFilterCreatedTo.value);
  state.serviceFilters.dateFrom = clean(els.servicesFilterDateFrom.value);
  state.serviceFilters.dateTo = clean(els.servicesFilterDateTo.value);
  state.serviceFilters.name = clean(els.servicesFilterName.value);
  renderServices();
}

function setServicesScreen(screen) {
  state.servicesScreen = screen === "resume" ? "resume" : "list";
  renderServices();
}

function renderServicesScreenTabs() {
  const isResume = state.servicesScreen === "resume";
  els.servicesTabList?.classList.toggle("active-tab", !isResume);
  els.servicesTabList?.classList.toggle("ghost", isResume);
  els.servicesTabResume?.classList.toggle("active-tab", isResume);
  els.servicesTabResume?.classList.toggle("ghost", !isResume);
  if (els.servicesPanelList) els.servicesPanelList.hidden = isResume;
  if (els.servicesPanelResume) els.servicesPanelResume.hidden = !isResume;
}

function formatServiceMonthLabel(monthKey) {
  const raw = clean(monthKey);
  if (!/^\d{4}-\d{2}$/.test(raw)) return raw || "-";
  const dt = new Date(`${raw}-01T00:00:00`);
  if (Number.isNaN(dt.getTime())) return raw;
  return new Intl.DateTimeFormat("en-GB", { month: "short", year: "numeric", timeZone: "Europe/Lisbon" }).format(dt);
}

function getServicesMonthlyResume() {
  const buckets = new Map();
  state.services.forEach((service) => {
    const date = clean(service.date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
    const monthKey = date.slice(0, 7);
    const current = buckets.get(monthKey) || { monthKey, totalServices: 0, sumPax: 0, sumPrice: 0 };
    current.totalServices += 1;
    current.sumPax += Math.max(0, Number(service.pax || 0));
    current.sumPrice += Math.max(0, Number(service.price || 0));
    buckets.set(monthKey, current);
  });
  return [...buckets.values()].sort((a, b) => clean(b.monthKey).localeCompare(clean(a.monthKey)));
}

function renderServicesResume() {
  if (!els.servicesResumeBody || !els.servicesResumeCount) return;
  const rows = getServicesMonthlyResume();
  els.servicesResumeCount.textContent = `${rows.length} month${rows.length === 1 ? "" : "s"}`;
  els.servicesResumeBody.innerHTML = "";
  if (!rows.length) {
    els.servicesResumeBody.innerHTML = '<tr><td colspan="4" class="empty">No services found.</td></tr>';
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(formatServiceMonthLabel(row.monthKey))}</td>
      <td>${escape(String(row.totalServices))}</td>
      <td>${escape(String(row.sumPax))}</td>
      <td>${escape(formatMoney(row.sumPrice))}</td>`;
    els.servicesResumeBody.appendChild(tr);
  });
}

function renderServices() {
  if (!els.servicesRows || !canApp("services")) return;
  renderServicesScreenTabs();
  if (state.servicesScreen === "resume") {
    renderServicesResume();
    return;
  }
  els.servicesShowActive.checked = !!state.serviceFilters.showActive;
  els.servicesFilterCreatedFrom.value = clean(state.serviceFilters.createdFrom);
  els.servicesFilterCreatedTo.value = clean(state.serviceFilters.createdTo);
  els.servicesFilterDateFrom.value = clean(state.serviceFilters.dateFrom);
  els.servicesFilterDateTo.value = clean(state.serviceFilters.dateTo);
  els.servicesFilterName.value = clean(state.serviceFilters.name);
  const rows = getFilteredServices();
  els.servicesCount.textContent = `${rows.length} service line${rows.length === 1 ? "" : "s"}`;
  els.servicesRows.innerHTML = "";
  if (els.servicesMobileCards) els.servicesMobileCards.innerHTML = "";
  if (!rows.length) {
    els.servicesRows.innerHTML = '<tr><td colspan="11" class="empty">No services found.</td></tr>';
    if (els.servicesMobileCards) {
      els.servicesMobileCards.innerHTML = '<div class="services-mobile-empty">No services found.</div>';
    }
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.className = `clickable-row${row.serviceId === state.serviceSelectedId ? " selected-row" : ""}${row.legType === "return" ? " service-return-row" : ""}`;
    tr.dataset.serviceId = row.serviceId;
    tr.title = clean(state.services.find((item) => item.id === row.serviceId)?.notes) || "-";
    tr.innerHTML = `<td>${escape(row.requestNumber || "-")}${row.legType === "return" ? ' <small>(return)</small>' : ""}</td>
      <td>${escape(row.serviceType)}</td>
      <td>${escape(row.customerName)}</td>
      <td>${renderServiceStatusCell(row)}</td>
      <td>${renderServiceDateCell(row.date)}</td>
      <td>${escape(clean(row.time) || "-")}</td>
      <td>${escape(String(row.pax || 0))}</td>
      <td>${escape(row.pickupLocation || "-")}</td>
      <td>${renderServiceFlightCell(row)}</td>
      <td>${escape(row.dropoffLocation || "-")}</td>
      <td>${escape(formatMoney(row.price))}</td>`;
    els.servicesRows.appendChild(tr);
    if (els.servicesMobileCards) {
      const card = document.createElement("article");
      card.className = `service-mobile-card${row.serviceId === state.serviceSelectedId ? " selected-card" : ""}${row.legType === "return" ? " service-return-row" : ""}`;
      card.dataset.serviceId = row.serviceId;
      card.title = clean(state.services.find((item) => item.id === row.serviceId)?.notes) || "-";
      card.innerHTML = `<div class="service-mobile-header">
          <div>
            <div class="service-mobile-request">${escape(row.requestNumber || "-")}${row.legType === "return" ? ' <small>(return)</small>' : ""}</div>
            <div class="service-mobile-type">${escape(row.serviceType)}</div>
          </div>
          ${renderServiceStatusCell(row)}
        </div>
        <div class="service-mobile-customer">${escape(row.customerName)}</div>
        <div class="service-mobile-grid">
          <div class="service-mobile-field">
            <small>Date</small>
            <div>${renderServiceDateCell(row.date)}</div>
          </div>
          <div class="service-mobile-field">
            <small>Time</small>
            <div>${escape(clean(row.time) || "-")}</div>
          </div>
          <div class="service-mobile-field">
            <small>Pax</small>
            <div>${escape(String(row.pax || 0))}</div>
          </div>
          <div class="service-mobile-field">
            <small>Pick Up</small>
            <div>${escape(row.pickupLocation || "-")}</div>
          </div>
          <div class="service-mobile-field">
            <small>Flight Nr</small>
            <div>${renderServiceFlightCell(row)}</div>
          </div>
          <div class="service-mobile-field">
            <small>Drop Off</small>
            <div>${escape(row.dropoffLocation || "-")}</div>
          </div>
        </div>
        <div class="service-mobile-price-row">
          <small>Price</small>
          <strong>${escape(formatMoney(row.price))}</strong>
        </div>`;
      els.servicesMobileCards.appendChild(card);
    }
  });
  refreshVisibleServiceStatuses();
}

function renderServiceSettings() {
  if (!els.servicesConfigsBody) return;
  const configs = serviceConfigs();
  if (els.servicesAutomaticEmailRecipients) {
    els.servicesAutomaticEmailRecipients.value = (state.serviceSettings?.automaticEmailRecipients || []).join("\n");
  }
  if (els.servicesLiveFlightStatusEnabled) {
    els.servicesLiveFlightStatusEnabled.checked = serviceLiveFlightStatusEnabled();
  }
  els.servicesConfigsBody.innerHTML = "";
  configs.forEach((config, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input data-service-setting-type="${index}" value="${escape(config.serviceType)}" /></td>
      <td><select data-service-setting-provider="${index}">${serviceProviderOptions(config.providerUserId)}</select></td>
      <td><input type="checkbox" data-service-setting-airport="${index}" ${config.airportTransfer ? "checked" : ""} /></td>
      <td><input type="checkbox" data-service-setting-return="${index}" ${config.hasReturn ? "checked" : ""} /></td>
      <td><input type="checkbox" data-service-setting-approved="${index}" ${config.approvedByDefault ? "checked" : ""} /></td>
      <td><select data-service-setting-price-mode="${index}">
        <option value="airport_matrix" ${config.priceMode === "airport_matrix" ? "selected" : ""}>Airport matrix</option>
        <option value="open" ${config.priceMode !== "airport_matrix" ? "selected" : ""}>Open</option>
      </select></td>`;
    els.servicesConfigsBody.appendChild(tr);
  });
  const airportConfig = serviceConfigByType("Airport Transfer") || configs[0] || sanitizeServiceConfigClient(DEFAULT_SERVICE_SETTINGS.serviceConfigs[0]);
  els.servicesPriceOneWay13.value = airportConfig?.priceMatrix?.oneWay?.["1-3"] ?? 0;
  els.servicesPriceOneWay47.value = airportConfig?.priceMatrix?.oneWay?.["4-7"] ?? 0;
  els.servicesPriceOneWay811.value = airportConfig?.priceMatrix?.oneWay?.["8-11"] ?? 0;
  els.servicesPriceOneWay1216.value = airportConfig?.priceMatrix?.oneWay?.["12-16"] ?? 0;
  els.servicesPriceReturn13.value = airportConfig?.priceMatrix?.returnTrip?.["1-3"] ?? 0;
  els.servicesPriceReturn47.value = airportConfig?.priceMatrix?.returnTrip?.["4-7"] ?? 0;
  els.servicesPriceReturn811.value = airportConfig?.priceMatrix?.returnTrip?.["8-11"] ?? 0;
  els.servicesPriceReturn1216.value = airportConfig?.priceMatrix?.returnTrip?.["12-16"] ?? 0;
  renderServiceSettingsTemplateEditor();
  renderServiceSettingsTab();
}

function onServiceSettingsInput() {
  state.serviceSettings.automaticEmailRecipients = parseEmailList(els.servicesAutomaticEmailRecipients?.value);
  state.serviceSettings.liveFlightStatusEnabled = !!els.servicesLiveFlightStatusEnabled?.checked;
  state.serviceSettings.serviceConfigs = serviceConfigs().map((config, index) => {
    const providerId = clean(els.servicesConfigsBody.querySelector(`[data-service-setting-provider="${index}"]`)?.value);
    const provider = state.serviceProviders.find((item) => clean(item.id) === providerId);
    return sanitizeServiceConfigClient({
      ...config,
      serviceType: clean(els.servicesConfigsBody.querySelector(`[data-service-setting-type="${index}"]`)?.value),
      providerUserId: providerId,
      providerEmail: provider?.email || config.providerEmail,
      airportTransfer: !!els.servicesConfigsBody.querySelector(`[data-service-setting-airport="${index}"]`)?.checked,
      hasReturn: !!els.servicesConfigsBody.querySelector(`[data-service-setting-return="${index}"]`)?.checked,
      approvedByDefault: !!els.servicesConfigsBody.querySelector(`[data-service-setting-approved="${index}"]`)?.checked,
      priceMode: clean(els.servicesConfigsBody.querySelector(`[data-service-setting-price-mode="${index}"]`)?.value),
      confirmationTemplate: config.confirmationTemplate,
    });
  });
  const airportConfig = serviceConfigByType("Airport Transfer") || state.serviceSettings.serviceConfigs[0];
  if (airportConfig) {
    airportConfig.priceMatrix.oneWay["1-3"] = Number(normalizeNumber(els.servicesPriceOneWay13.value) || 0);
    airportConfig.priceMatrix.oneWay["4-7"] = Number(normalizeNumber(els.servicesPriceOneWay47.value) || 0);
    airportConfig.priceMatrix.oneWay["8-11"] = Number(normalizeNumber(els.servicesPriceOneWay811.value) || 0);
    airportConfig.priceMatrix.oneWay["12-16"] = Number(normalizeNumber(els.servicesPriceOneWay1216.value) || 0);
    airportConfig.priceMatrix.returnTrip["1-3"] = Number(normalizeNumber(els.servicesPriceReturn13.value) || 0);
    airportConfig.priceMatrix.returnTrip["4-7"] = Number(normalizeNumber(els.servicesPriceReturn47.value) || 0);
    airportConfig.priceMatrix.returnTrip["8-11"] = Number(normalizeNumber(els.servicesPriceReturn811.value) || 0);
    airportConfig.priceMatrix.returnTrip["12-16"] = Number(normalizeNumber(els.servicesPriceReturn1216.value) || 0);
  }
  const selectedTemplateConfig = currentServiceSettingsTemplateConfig();
  if (selectedTemplateConfig && els.servicesConfirmationTemplate) {
    selectedTemplateConfig.confirmationTemplate = els.servicesConfirmationTemplate.value;
  }
  state.serviceFlightStatuses.cache = {};
  state.serviceFlightStatuses.initialized = false;
  if (state.currentView === "services") renderServices();
  renderServiceSettingsTemplatePreview();
}

async function saveServiceSettings() {
  onServiceSettingsInput();
  state.serviceSettings = sanitizeServiceSettingsClient(state.serviceSettings);
  renderServiceSettings();
  try {
    const result = await api("/api/service-settings", { method: "PUT", body: { settings: state.serviceSettings } });
    state.serviceSettings = sanitizeServiceSettingsClient(result.settings);
    state.serviceProviders = Array.isArray(result.providers) ? result.providers : state.serviceProviders;
    state.serviceSettingsLoaded = true;
    renderServiceSettings();
    setServicesSettingsStatus("Services configuration saved.");
    showToast("Services configuration saved.", "success");
  } catch (e) {
    setServicesSettingsStatus(`Save failed: ${e.message}`);
    showToast(`Services configuration save failed: ${e.message}`, "error");
  }
}

function renderServiceSettingsTab() {
  const isConfig = state.serviceSettingsTab !== "confirmation";
  els.servicesSettingsConfigTab.classList.toggle("active-tab", isConfig);
  els.servicesSettingsConfigTab.classList.toggle("ghost", !isConfig);
  els.servicesSettingsConfirmationTab.classList.toggle("active-tab", !isConfig);
  els.servicesSettingsConfirmationTab.classList.toggle("ghost", isConfig);
  els.servicesSettingsConfigPanel.hidden = !isConfig;
  els.servicesSettingsConfirmationPanel.hidden = isConfig;
}

function setServiceSettingsTab(tab) {
  state.serviceSettingsTab = tab === "confirmation" ? "confirmation" : "config";
  renderServiceSettingsTab();
}

function renderServiceSettingsTemplateEditor() {
  if (!els.servicesTemplateServiceType || !els.servicesConfirmationTemplate) return;
  const configs = serviceConfigs();
  const selectedId = ensureServiceSettingsTemplateType();
  els.servicesTemplateServiceType.innerHTML = configs.length
    ? configs.map((config) => `<option value="${escape(config.id)}">${escape(config.serviceType)}</option>`).join("")
    : '<option value="">No service types</option>';
  els.servicesTemplateServiceType.value = selectedId;
  if (els.servicesTemplateLanguage) els.servicesTemplateLanguage.value = normalizeServiceConfirmationLanguage(state.serviceSettingsTemplateLanguage);
  const config = currentServiceSettingsTemplateConfig();
  els.servicesConfirmationTemplate.value = config ? serviceConfirmationTemplate(config) : "";
  renderServiceSettingsTemplatePreview();
}

function renderServiceSettingsTemplatePreview() {
  if (!els.servicesConfirmationTemplatePreview) return;
  const config = currentServiceSettingsTemplateConfig();
  if (!config) {
    els.servicesConfirmationTemplatePreview.innerHTML = "<h4>Preview</h4><p>No service type available.</p>";
    return;
  }
  const draft = serviceTemplatePreviewDraft(config);
  const language = normalizeServiceConfirmationLanguage(state.serviceSettingsTemplateLanguage);
  const template = language === "en"
    ? (clean(els.servicesConfirmationTemplate.value) || serviceConfirmationTemplate(config, "en"))
    : serviceConfirmationTemplate(config, language);
  els.servicesConfirmationTemplatePreview.innerHTML = `<h4>Preview</h4>${serviceConfirmationEmailHtmlFromTemplate(draft, template, config, language)}`;
}

function onServiceSettingsTemplateChange() {
  onServiceSettingsInput();
  state.serviceSettingsTemplateType = clean(els.servicesTemplateServiceType.value);
  renderServiceSettingsTemplateEditor();
}

function resetServiceDraft() {
  state.serviceSelectedId = "";
  resetServiceDraftPredictionState({ keepCache: true });
  const draft = emptyServiceDraft();
  const firstConfig = serviceConfigs()[0];
  if (firstConfig) draft.serviceType = firstConfig.serviceType;
  state.serviceDraft = draft;
  applyServiceConfigToDraft({ forcePrice: true, fromTypeChange: true });
}

function applyServiceConfigToDraft({ forcePrice = false, fromTypeChange = false } = {}) {
  const draft = state.serviceDraft;
  const config = serviceConfigByType(draft.serviceType);
  if (!config) return;
  const showFlight = serviceUsesFlightFields(config, draft.serviceType);
  draft.providerUserId = clean(config.providerUserId);
  draft.providerEmail = clean(config.providerEmail);
  if (fromTypeChange && !draft.id) draft.status = config.approvedByDefault ? "Approved" : "Submitted";
  if (!config.hasReturn) draft.hasReturn = false;
  if (!showFlight) {
    draft.flightNumber = "";
    draft.returnFlight = "";
  }
  if (draft.hasReturn) {
    if (!clean(draft.returnPickup)) draft.returnPickup = clean(draft.dropoffLocation);
    if (!clean(draft.returnDropoff)) draft.returnDropoff = clean(draft.pickupLocation);
  } else {
    draft.returnPickup = "";
    draft.returnDropoff = "";
    draft.returnDate = "";
    draft.returnTime = "";
    draft.returnFlight = "";
  }
  const autoPrice = serviceComputedPrice(config, draft.pax, draft.hasReturn);
  if ((forcePrice || !draft.priceManual) && autoPrice !== null) draft.price = autoPrice;
}

function renderServiceDraft() {
  const draft = state.serviceDraft;
  const config = serviceConfigByType(draft.serviceType);
  const hasReturnAvailable = !!config?.hasReturn;
  const showFlight = serviceUsesFlightFields(config, draft.serviceType);
  renderServiceEditorTab();
  state.serviceConfirmationLanguage = normalizeServiceConfirmationLanguage(draft.language || state.serviceConfirmationLanguage);
  if (els.serviceRequestNumberLabel) els.serviceRequestNumberLabel.textContent = clean(draft.requestNumber) ? `#${clean(draft.requestNumber)}` : "";
  els.serviceType.innerHTML = serviceConfigs().map((item) => option(item.serviceType, draft.serviceType)).join("");
  els.serviceType.value = draft.serviceType || clean(serviceConfigs()[0]?.serviceType);
  els.serviceStatus.value = normalizeServiceStatus(draft.status);
  els.serviceCustomerName.value = draft.customerName;
  els.serviceCustomerEmail.value = draft.customerEmail;
  els.serviceCustomerPhone.value = draft.customerPhone;
  renderServicePhoneFlag();
  els.servicePax.value = String(draft.pax || 1);
  els.serviceDate.value = formatServiceDateInput(draft.date);
  els.serviceTime.value = draft.time;
  els.servicePickupLocation.value = draft.pickupLocation;
  els.serviceDropoffLocation.value = draft.dropoffLocation;
  els.serviceFlightNumber.value = draft.flightNumber;
  els.serviceHasReturn.value = draft.hasReturn ? "true" : "false";
  els.serviceHasReturn.disabled = !hasReturnAvailable;
  els.servicePrice.value = draft.price ?? 0;
  els.serviceProviderEmail.textContent = draft.providerEmail ? `Service provider: ${draft.providerEmail}` : "";
  els.serviceNotes.value = draft.notes;
  els.serviceReturnPickup.value = draft.returnPickup;
  els.serviceReturnDropoff.value = draft.returnDropoff;
  els.serviceReturnDate.value = formatServiceDateInput(draft.returnDate);
  els.serviceReturnTime.value = draft.returnTime;
  els.serviceReturnFlight.value = draft.returnFlight;
  if (els.serviceConfirmationLanguage) els.serviceConfirmationLanguage.value = state.serviceConfirmationLanguage;
  els.serviceFlightField.hidden = false;
  els.serviceFlightNumber.disabled = !showFlight;
  els.serviceReturnFields.hidden = !(hasReturnAvailable && draft.hasReturn);
  els.serviceReturnFlightField.hidden = false;
  els.serviceReturnFlight.disabled = !showFlight;
  els.serviceDelete.hidden = true;
  renderServiceAuditHistory();
  renderServiceConfirmationPreview();
  syncServiceDatePickers();
  queueServiceDraftPredictionRefresh();
}

function renderServiceAuditHistory() {
  const audit = normalizeServiceAudit(state.serviceDraft.audit);
  if (!audit.length) {
    els.serviceAuditHistory.classList.add("empty");
    els.serviceAuditHistory.innerHTML = "No saved changes yet.";
    return;
  }
  els.serviceAuditHistory.classList.remove("empty");
  els.serviceAuditHistory.innerHTML = audit
    .slice()
    .reverse()
    .map((item) => `<article><strong>${escape(item.action)}</strong><span>${escape(formatDateTimeShort(item.at))}</span><span>${escape(item.user || "-")}</span><p>${escape(item.summary || "-")}</p></article>`)
    .join("");
}

function openServiceModal() {
  state.serviceEditorTab = "details";
  state.serviceConfirmationLanguage = normalizeServiceConfirmationLanguage(state.serviceDraft.language || state.serviceConfirmationLanguage);
  els.serviceEditorModal.hidden = false;
  document.body.classList.add("modal-open");
  renderServiceDraft();
}

function closeServiceModal() {
  els.serviceEditorModal.hidden = true;
  document.body.classList.remove("modal-open");
  resetServiceDraftPredictionState({ keepCache: true });
  state.serviceSelectedId = "";
  syncAppRoute();
}

async function openServiceById(serviceId, { updateUrl = true } = {}) {
  const needle = clean(serviceId);
  const service = state.services.find((item) => item.id === needle || item.requestNumber === needle);
  if (!service) return false;
  try {
    setServicesStatus("Loading latest service...");
    const result = await api(`/api/services?id=${encodeURIComponent(service.id)}`);
    const latest = result?.row ? mapServiceRow(result.row) : null;
    if (!latest) throw new Error("Service not found.");
    const index = state.services.findIndex((item) => item.id === latest.id);
    if (index === -1) state.services.push(latest);
    else state.services.splice(index, 1, latest);
    state.serviceSelectedId = latest.id;
    state.serviceDraft = { ...clone(latest), language: normalizeServiceConfirmationLanguage(latest.language), priceManual: false };
    state.pendingServiceDeepLinkId = "";
    if (updateUrl) syncAppRoute();
    openServiceModal();
    setServicesStatus("Latest service loaded.");
    return true;
  } catch (e) {
    setServicesStatus(`Could not load latest service: ${e.message}`);
    showToast(`Could not load latest service: ${e.message}`, "error");
    return false;
  }
}

async function tryOpenDeepLinkedService() {
  const serviceId = clean(state.pendingServiceDeepLinkId);
  if (!serviceId || !state.services.length || !canApp("services")) return;
  if (!(await openServiceById(serviceId, { updateUrl: true }))) {
    state.pendingServiceDeepLinkId = "";
  }
}

async function onServiceRowClick(event) {
  if (event.target.closest("[data-inline-service-status]")) return;
  const row = event.target.closest("[data-service-id]");
  if (!row) return;
  await openServiceById(clean(row.dataset.serviceId), { updateUrl: true });
}

function buildServicePayload(draft, previous) {
  return {
    serviceType: draft.serviceType,
    customerName: draft.customerName,
    customerEmail: draft.customerEmail,
    customerPhone: draft.customerPhone,
    pax: draft.pax,
    notes: draft.notes,
    serviceDate: draft.date,
    serviceTime: draft.time,
    pickupLocation: draft.pickupLocation,
    dropoffLocation: draft.dropoffLocation,
    flightNumber: draft.flightNumber,
    hasReturn: draft.hasReturn,
    returnPickupLocation: draft.returnPickup,
    returnDropoffLocation: draft.returnDropoff,
    returnDate: draft.returnDate,
    returnTime: draft.returnTime,
    returnFlightNumber: draft.returnFlight,
    price: Number(draft.price || 0),
    status: draft.status,
    providerUserId: draft.providerUserId || null,
    providerEmail: draft.providerEmail,
    auditLog: appendServiceAudit(draft, previous),
  };
}

async function updateServiceStatusInline(serviceId, status) {
  const previous = state.services.find((item) => item.id === clean(serviceId));
  if (!previous) return;
  const nextStatus = normalizeServiceStatus(status);
  if (normalizeServiceStatus(previous.status) === nextStatus) return;
  const draft = { ...clone(previous), status: nextStatus };
  state.serviceInlineStatusSaving[serviceId] = true;
  renderServices();
  try {
    const result = await api(`/api/services?id=${encodeURIComponent(previous.id)}`, {
      method: "PUT",
      body: buildServicePayload(draft, previous),
    });
    const row = result?.row ? mapServiceRow(result.row) : null;
    if (row) {
      const index = state.services.findIndex((item) => item.id === row.id);
      if (index === -1) state.services.push(row);
      else state.services.splice(index, 1, row);
    }
    await loadServices({ silent: true, throwOnError: true });
    state.servicesLoaded = true;
    const baseStatus = "Service status updated.";
    const emailWarning = clean(result?.emailWarning);
    setServicesStatus(emailWarning ? `${baseStatus} Email warning: ${emailWarning}` : baseStatus);
    setServicesDbStatus(`Loaded ${state.services.length} service request${state.services.length === 1 ? "" : "s"}.`);
    showToast(emailWarning ? `${baseStatus} Email warning: ${emailWarning}` : baseStatus, emailWarning ? "warning" : "success");
  } catch (e) {
    setServicesStatus(`Status update failed: ${e.message}`);
    showToast(`Status update failed: ${e.message}`, "error");
  } finally {
    delete state.serviceInlineStatusSaving[serviceId];
    renderServices();
  }
}

function onInlineServiceStatusChange(event) {
  const select = event.target.closest("[data-inline-service-status]");
  if (!select) return;
  event.stopPropagation();
  updateServiceStatusInline(clean(select.dataset.inlineServiceStatus), clean(select.value));
}

function onServiceDatePickerInput(event) {
  if (event.target === els.serviceDatePicker) {
    state.serviceDraft.date = clean(els.serviceDatePicker.value);
    els.serviceDate.value = formatServiceDateInput(state.serviceDraft.date);
    if (state.serviceDraft.returnDate && state.serviceDraft.returnDate < state.serviceDraft.date) {
      state.serviceDraft.returnDate = state.serviceDraft.date;
      els.serviceReturnDate.value = formatServiceDateInput(state.serviceDraft.returnDate);
    }
  }
  if (event.target === els.serviceReturnDatePicker) {
    state.serviceDraft.returnDate = clean(els.serviceReturnDatePicker.value);
    els.serviceReturnDate.value = formatServiceDateInput(state.serviceDraft.returnDate);
  }
  renderServiceDraft();
}

function renderServiceEditorTab() {
  const isDetails = state.serviceEditorTab !== "confirmation";
  if (els.serviceTabDetails) {
    els.serviceTabDetails.classList.toggle("active-tab", isDetails);
    els.serviceTabDetails.classList.toggle("ghost", !isDetails);
  }
  if (els.serviceTabConfirmation) {
    els.serviceTabConfirmation.classList.toggle("active-tab", !isDetails);
    els.serviceTabConfirmation.classList.toggle("ghost", isDetails);
  }
  if (els.serviceDetailsPanel) els.serviceDetailsPanel.hidden = !isDetails;
  if (els.serviceConfirmationPanel) els.serviceConfirmationPanel.hidden = isDetails;
}

function setServiceEditorTab(tab) {
  state.serviceEditorTab = tab === "confirmation" ? "confirmation" : "details";
  renderServiceEditorTab();
  renderServiceConfirmationPreview();
}

function renderServiceConfirmationPreview() {
  if (!els.serviceConfirmationPreview) return;
  const config = serviceConfigByType(state.serviceDraft.serviceType);
  if (!config) {
    els.serviceConfirmationPreview.innerHTML = "Select a service type to preview the confirmation text.";
    return;
  }
  state.serviceDraft.language = normalizeServiceConfirmationLanguage(state.serviceDraft.language || state.serviceConfirmationLanguage);
  els.serviceConfirmationPreview.innerHTML = serviceConfirmationEmailHtml(state.serviceDraft);
}

async function copyServiceConfirmationText() {
  const text = serviceConfirmationEmailText(state.serviceDraft);
  const html = serviceConfirmationEmailHtml(state.serviceDraft);
  try {
    if (window.ClipboardItem) {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" }),
        }),
      ]);
      setServicesStatus("Formatted confirmation email copied.");
      showToast("Formatted confirmation email copied.", "success");
    } else {
      await navigator.clipboard.writeText(text);
      setServicesStatus("Plain confirmation email text copied.");
      showToast("Plain confirmation email text copied.", "success");
    }
  } catch (e) {
    setServicesStatus("Could not copy automatically. Select the preview and copy it manually.");
  }
}

function onServiceDraftInput(event) {
  const draft = state.serviceDraft;
  const previousType = draft.serviceType;
  draft.serviceType = clean(els.serviceType.value);
  draft.status = normalizeServiceStatus(els.serviceStatus.value);
  draft.customerName = draftText(els.serviceCustomerName.value);
  draft.customerEmail = clean(els.serviceCustomerEmail.value).toLowerCase();
  draft.customerPhone = draftText(els.serviceCustomerPhone.value);
  draft.pax = Math.max(1, Math.min(60, Math.round(Number(normalizeNumber(els.servicePax.value) || 1))));
  draft.date = parseServiceDateInput(els.serviceDate.value);
  draft.time = clean(els.serviceTime.value);
  draft.pickupLocation = draftText(els.servicePickupLocation.value);
  draft.dropoffLocation = draftText(els.serviceDropoffLocation.value);
  draft.flightNumber = clean(els.serviceFlightNumber.value);
  draft.hasReturn = normalizeServiceBool(els.serviceHasReturn.value);
  draft.notes = draftText(els.serviceNotes.value);
  draft.returnPickup = draftText(els.serviceReturnPickup.value);
  draft.returnDropoff = draftText(els.serviceReturnDropoff.value);
  draft.returnDate = parseServiceDateInput(els.serviceReturnDate.value);
  draft.returnTime = clean(els.serviceReturnTime.value);
  draft.returnFlight = clean(els.serviceReturnFlight.value);
  if (event?.target === els.servicePrice) {
    draft.price = Number(normalizeNumber(els.servicePrice.value) || 0);
    draft.priceManual = true;
  } else {
    draft.price = Number(normalizeNumber(els.servicePrice.value) || draft.price || 0);
  }
  if (draft.hasReturn) {
    if (event?.target === els.serviceDropoffLocation && !clean(draft.returnPickup)) draft.returnPickup = draft.dropoffLocation;
    if (event?.target === els.servicePickupLocation && !clean(draft.returnDropoff)) draft.returnDropoff = draft.pickupLocation;
    if (draft.date && draft.returnDate && draft.returnDate < draft.date) draft.returnDate = draft.date;
  }
  syncServiceDatePickers();
  applyServiceConfigToDraft({ forcePrice: false, fromTypeChange: previousType !== draft.serviceType });
  renderServiceDraft();
}

function serviceAuditSummary(draft, previous) {
  if (!previous?.id) {
    return `Created ${clean(draft.serviceType) || "service"} for ${clean(draft.customerName) || "customer"} on ${formatDateOnly(draft.date)} ${clean(draft.time)} for ${draft.pax || 0} pax (${formatMoney(draft.price)}).`;
  }
  const formatValue = (value, kind = "text") => {
    if (kind === "money") return formatMoney(value || 0);
    if (kind === "bool") return value ? "Yes" : "No";
    if (kind === "date") return clean(value) ? formatDateOnly(value) : "-";
    if (kind === "time") return clean(value) || "-";
    if (kind === "number") return clean(value) ? String(value) : "0";
    return clean(value) || "-";
  };
  const pushChange = (changes, label, before, after, kind = "text") => {
    const beforeText = formatValue(before, kind);
    const afterText = formatValue(after, kind);
    if (beforeText !== afterText) changes.push(`${label}: ${beforeText} -> ${afterText}`);
  };
  const changes = [];
  pushChange(changes, "Service Type", previous.serviceType, draft.serviceType);
  pushChange(changes, "Status", previous.status, draft.status);
  pushChange(changes, "Customer Name", previous.customerName, draft.customerName);
  pushChange(changes, "Customer Email", previous.customerEmail, draft.customerEmail);
  pushChange(changes, "Customer Phone", previous.customerPhone, draft.customerPhone);
  pushChange(changes, "Date", previous.date, draft.date, "date");
  pushChange(changes, "Time", previous.time, draft.time, "time");
  pushChange(changes, "Pax", previous.pax, draft.pax, "number");
  pushChange(changes, "Pickup Location", previous.pickupLocation, draft.pickupLocation);
  pushChange(changes, "Drop Off Location", previous.dropoffLocation, draft.dropoffLocation);
  pushChange(changes, "Flight Number", previous.flightNumber, draft.flightNumber);
  pushChange(changes, "Return?", previous.hasReturn, draft.hasReturn, "bool");
  pushChange(changes, "Return Pickup", previous.returnPickup, draft.returnPickup);
  pushChange(changes, "Return Drop Off", previous.returnDropoff, draft.returnDropoff);
  pushChange(changes, "Return Date", previous.returnDate, draft.returnDate, "date");
  pushChange(changes, "Return Time", previous.returnTime, draft.returnTime, "time");
  pushChange(changes, "Return Flight", previous.returnFlight, draft.returnFlight);
  if (Math.abs(Number(previous.price || 0) - Number(draft.price || 0)) >= 0.01) pushChange(changes, "Price", previous.price, draft.price, "money");
  pushChange(changes, "Provider", previous.providerEmail, draft.providerEmail);
  pushChange(changes, "Notes", previous.notes, draft.notes);
  return changes.length ? changes.join("; ") : "Saved without major field changes.";
}

function appendServiceAudit(draft, previous) {
  const current = normalizeServiceAudit(draft.audit);
  const action = previous?.id ? "Updated service" : "Created service";
  return current.concat([{
    at: new Date().toISOString(),
    action,
    user: clean(state.user?.email) || "Unknown user",
    summary: serviceAuditSummary(draft, previous),
  }]).slice(-20);
}

async function saveService() {
  const draft = state.serviceDraft;
  if (!clean(draft.serviceType)) return setServicesStatus("Service Type is required.");
  if (!clean(draft.customerName)) return setServicesStatus("Customer Name is required.");
  if (!clean(draft.customerPhone)) return setServicesStatus("Customer Phone is required.");
  if (!clean(draft.date)) return setServicesStatus("Date is required.");
  if (!clean(draft.time)) return setServicesStatus("Time is required.");
  if (clean(draft.customerPhone) && !isValidInternationalPhone(draft.customerPhone)) {
    return setServicesStatus("Customer Phone must include country code, for example +351 912 345 678.");
  }
  if (draft.hasReturn && clean(draft.returnDate) && clean(draft.date) && clean(draft.returnDate) < clean(draft.date)) {
    return setServicesStatus("Return date cannot be before the main service date.");
  }
  const previous = draft.id ? state.services.find((item) => item.id === draft.id) : null;
  const payload = buildServicePayload(draft, previous);
  try {
    const result = previous?.id
      ? await api(`/api/services?id=${encodeURIComponent(previous.id)}`, { method: "PUT", body: payload })
      : await api("/api/services", { method: "POST", body: payload });
    const row = result?.row ? mapServiceRow(result.row) : null;
    if (row) {
      const index = state.services.findIndex((item) => item.id === row.id);
      if (index === -1) state.services.push(row);
      else state.services.splice(index, 1, row);
      state.serviceSelectedId = row.id;
      state.serviceDraft = { ...clone(row), priceManual: false };
    }
    await loadServices({ silent: true, throwOnError: true });
    state.servicesLoaded = true;
    closeServiceModal();
    const baseStatus = previous?.id ? "Service updated." : "Service created.";
    const emailWarning = clean(result?.emailWarning);
    setServicesStatus(emailWarning ? `${baseStatus} Email warning: ${emailWarning}` : baseStatus);
    setServicesDbStatus(`Loaded ${state.services.length} service request${state.services.length === 1 ? "" : "s"}.`);
    showToast(emailWarning ? `${baseStatus} Email warning: ${emailWarning}` : baseStatus, emailWarning ? "warning" : "success");
  } catch (e) {
    setServicesStatus(`Save failed: ${e.message}`);
    showToast(`Service save failed: ${e.message}`, "error");
  }
}

async function deleteService() {
  const id = clean(state.serviceDraft.id);
  if (!id) return;
  if (!window.confirm("Delete this service request?")) return;
  try {
    await api(`/api/services?id=${encodeURIComponent(id)}`, { method: "DELETE" });
    state.services = state.services.filter((item) => item.id !== id);
    resetServiceDraft();
    closeServiceModal();
    renderServices();
    setServicesStatus("Service deleted.");
    setServicesDbStatus(`Loaded ${state.services.length} service request${state.services.length === 1 ? "" : "s"}.`);
    showToast("Service deleted.", "success");
  } catch (e) {
    setServicesStatus(`Delete failed: ${e.message}`);
    showToast(`Service delete failed: ${e.message}`, "error");
  }
}

function exportServicesToExcel() {
  const rows = getFilteredServices();
  const headers = ["Request #", "Service Type", "Customer", "Date", "Time", "Pax", "Pick Up", "Flight Nr", "Drop Off", "Price"];
  const body = rows.map((row) => [
    row.requestNumber || "-",
    row.serviceType,
    row.customerName,
    formatDateOnly(row.date),
    row.time || "-",
    row.pax || 0,
    row.pickupLocation || "-",
    row.flightNumber || "-",
    row.dropoffLocation || "-",
    formatMoney(row.price),
  ]);
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><table border="1"><thead><tr>${headers.map((header) => `<th>${escape(header)}</th>`).join("")}</tr></thead><tbody>${body.map((cols) => `<tr>${cols.map((value) => `<td>${escape(value)}</td>`).join("")}</tr>`).join("")}</tbody></table></body></html>`;
  downloadBlob(`services_${formatDate(new Date())}.xls`, html, "application/vnd.ms-excel;charset=utf-8;");
  showToast(`Exported ${rows.length} active services to Excel.`, "success");
}

function normalizeShoppingWeekdaysClient(value) {
  const source = Array.isArray(value) ? value : [];
  const seen = new Set();
  return source
    .map((item) => clean(item).toLowerCase())
    .filter((item) => SHOPPING_WEEKDAY_OPTIONS.some((entry) => entry.key === item))
    .filter((item) => {
      if (seen.has(item)) return false;
      seen.add(item);
      return true;
    });
}

function normalizeShoppingCategoryClient(value) {
  const raw = clean(value).toLowerCase();
  if (raw === "breakfast" || raw === "pequeno almoço" || raw === "pequeno almoco") return "Breakfast";
  if (raw === "cleaning" || raw === "limpeza") return "Cleaning";
  if (raw === "sales" || raw === "vendas") return "Sales";
  if (raw === "activities" || raw === "atividades" || raw === "actividades") return "Activities";
  if (raw === "other" || raw === "outros") return "Other";
  if (raw === "tapas" || raw.includes("terça") || raw.includes("terca")) return "Tapas";
  if (raw === "utensils" || raw === "utensilios" || raw === "utensílios") return "Utensils";
  return SHOPPING_CATEGORY_OPTIONS.includes(clean(value)) ? clean(value) : "Other";
}

function normalizeShoppingStoredClient(value) {
  const raw = clean(value);
  if (!raw) return "";
  const exact = SHOPPING_STORED_OPTIONS.find((option) => option === raw);
  if (exact) return exact;
  const normalized = raw.toLowerCase();
  return SHOPPING_STORED_OPTIONS.find((option) => option.toLowerCase() === normalized) || raw;
}

function normalizeShoppingColorClient(value, fallback = "#F3E7DB") {
  const raw = clean(value).toUpperCase();
  return /^#[0-9A-F]{6}$/.test(raw) ? raw : fallback;
}

function parseShoppingBoolClient(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const raw = clean(value).toLowerCase();
  if (!raw) return false;
  return ["true", "1", "yes", "y", "sim", "on"].includes(raw);
}

function sanitizeShoppingSettingItemClient(item = {}, index = 0) {
  const label = clean(item.item || item.name);
  return {
    id: clean(item.id) || `shopping-item-${index + 1}`,
    category: normalizeShoppingCategoryClient(item.category),
    item: label,
    supplier: clean(item.supplier || item.suppliers),
    stored: normalizeShoppingStoredClient(item.stored),
    quantityRequired: parseShoppingBoolClient(item.quantityRequired ?? item.quantity_required ?? item.mandatoryExistingQuantity ?? item.mandatory_existing_quantity),
  };
}

function normalizeShoppingSettingsClient(settings = {}) {
  const sourceItems = Array.isArray(settings.items) ? settings.items : [];
  const sourceColors =
    settings.categoryColors && typeof settings.categoryColors === "object"
      ? settings.categoryColors
      : settings.category_colors && typeof settings.category_colors === "object"
        ? settings.category_colors
        : {};
  return {
    mandatoryWeekdays: normalizeShoppingWeekdaysClient(settings.mandatoryWeekdays || settings.mandatory_weekdays),
    emailRecipients: parseEmailList(settings.emailRecipients || settings.email_recipients),
    categoryColors: SHOPPING_CATEGORY_OPTIONS.reduce((acc, category) => {
      acc[category] = normalizeShoppingColorClient(sourceColors[category], DEFAULT_SHOPPING_CATEGORY_COLORS[category]);
      return acc;
    }, {}),
    items: sourceItems.map((item, index) => sanitizeShoppingSettingItemClient(item, index)).filter((item) => item.item),
  };
}

function normalizeShoppingOrderItemClient(item = {}) {
  return {
    id: clean(item.id),
    category: normalizeShoppingCategoryClient(item.category),
    item: clean(item.item),
    supplier: clean(item.supplier),
    stored: normalizeShoppingStoredClient(item.stored),
    quantityRequired: parseShoppingBoolClient(item.quantityRequired ?? item.quantity_required),
    existingQuantity: clean(item.existingQuantity ?? item.existing_quantity),
    order: parseShoppingBoolClient(item.order),
  };
}

function applyShoppingSettingsToOrderItemsClient(items, settingsItems = []) {
  const configMap = new Map((Array.isArray(settingsItems) ? settingsItems : []).map((item) => [clean(item.id), item]));
  const seen = new Set();
  const normalizedItems = (Array.isArray(items) ? items : []).map((item) => {
    const config = configMap.get(clean(item.id)) || null;
    const merged = !config ? item : {
      ...item,
      category: clean(config.category) || item.category,
      item: clean(config.item) || item.item,
      supplier: clean(config.supplier) || item.supplier,
      stored: clean(config.stored) || item.stored,
      quantityRequired: parseShoppingBoolClient(config.quantityRequired),
    };
    if (clean(merged.id)) seen.add(clean(merged.id));
    return merged;
  });
  (Array.isArray(settingsItems) ? settingsItems : []).forEach((config) => {
    const id = clean(config.id);
    if (!id || seen.has(id)) return;
    normalizedItems.push({
      id,
      category: clean(config.category),
      item: clean(config.item),
      supplier: clean(config.supplier),
      stored: clean(config.stored),
      quantityRequired: parseShoppingBoolClient(config.quantityRequired),
      existingQuantity: "",
      order: false,
    });
  });
  return normalizedItems;
}

function normalizeShoppingOrderClient(order, settingsItems = []) {
  if (!order) return null;
  const items = applyShoppingSettingsToOrderItemsClient(
    (Array.isArray(order.items) ? order.items : []).map(normalizeShoppingOrderItemClient).filter((item) => item.item),
    settingsItems
  );
  return {
    id: clean(order.id),
    orderNumber: Number(order.orderNumber || order.order_number || 0) || 0,
    status: clean(order.status).toLowerCase() === "submitted" ? "submitted" : "open",
    createdAt: clean(order.createdAt || order.created_at),
    updatedAt: clean(order.updatedAt || order.updated_at),
    submittedAt: clean(order.submittedAt || order.submitted_at),
    submittedByName: clean(order.submittedByName || order.submitted_by_name),
    submittedByUserEmail: clean(order.submittedByUserEmail || order.submitted_by_user_email).toLowerCase(),
    notes: clean(order.notes),
    reopenedFromId: clean(order.reopenedFromId || order.reopened_from_id),
    items,
    orderedCount: items.filter((item) => item.order).length,
  };
}

function normalizeBakeryBaseClient(value) {
  const raw = clean(value).toLowerCase().replace(/\s+/g, "-");
  if (["base-baixa", "baixa", "low"].includes(raw)) return "base-baixa";
  if (["base-alta", "alta", "high"].includes(raw)) return "base-alta";
  return "base-media";
}

function sanitizeBakeryBreadTableRowClient(row = {}, index = 0) {
  return {
    guests: Math.max(1, Number.parseInt(row.guests ?? index + 1, 10) || index + 1),
    baseBaixa: Math.max(0, Number.parseInt(row.baseBaixa ?? row.base_baixa ?? 0, 10) || 0),
    baseMedia: Math.max(0, Number.parseInt(row.baseMedia ?? row.base_media ?? 0, 10) || 0),
    baseAlta: Math.max(0, Number.parseInt(row.baseAlta ?? row.base_alta ?? 0, 10) || 0),
  };
}

function sanitizeBakeryBreadTypeClient(item = {}, index = 0) {
  return {
    id: clean(item.id) || `bread-type-${index + 1}`,
    name: clean(item.name || item.type),
    percentage: Math.max(0, Number(item.percentage || 0) || 0),
  };
}

function normalizeBakeryEmailProviderClient(value) {
  return clean(value).toLowerCase() === "smtp" ? "smtp" : "resend";
}

function normalizeBakerySecureClient(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  return ["true", "1", "yes", "sim", "on"].includes(clean(value).toLowerCase());
}

function normalizeBakeryEmailConfigClient(config = {}) {
  return {
    provider: normalizeBakeryEmailProviderClient(config.provider),
    smtpHost: clean(config.smtpHost || config.smtp_host) || DEFAULT_BAKERY_SETTINGS.emailConfig.smtpHost,
    smtpPort: Math.max(1, Number.parseInt(config.smtpPort ?? config.smtp_port ?? DEFAULT_BAKERY_SETTINGS.emailConfig.smtpPort, 10) || DEFAULT_BAKERY_SETTINGS.emailConfig.smtpPort),
    smtpSecure: config.smtpSecure === undefined && config.smtp_secure === undefined
      ? !!DEFAULT_BAKERY_SETTINGS.emailConfig.smtpSecure
      : normalizeBakerySecureClient(config.smtpSecure ?? config.smtp_secure),
    smtpUser: clean(config.smtpUser || config.smtp_user).toLowerCase(),
    smtpPassword: String(config.smtpPassword ?? config.smtp_password ?? ""),
    fromEmail: clean(config.fromEmail || config.from_email).toLowerCase(),
    fromName: clean(config.fromName || config.from_name) || DEFAULT_BAKERY_SETTINGS.emailConfig.fromName,
  };
}

function normalizeBakerySettingsClient(settings = {}) {
  return {
    selectedBase: normalizeBakeryBaseClient(settings.selectedBase || settings.selected_base),
    hostelCapacity: Math.max(1, Number.parseInt(settings.hostelCapacity ?? settings.hostel_capacity ?? 83, 10) || 83),
    emailRecipients: parseEmailList(settings.emailRecipients || settings.email_recipients),
    emailConfig: normalizeBakeryEmailConfigClient(settings.emailConfig || settings.email_config),
    breadTable: (Array.isArray(settings.breadTable) ? settings.breadTable : []).map(sanitizeBakeryBreadTableRowClient).sort((a, b) => a.guests - b.guests),
    breadTypes: (Array.isArray(settings.breadTypes) ? settings.breadTypes : []).map(sanitizeBakeryBreadTypeClient).filter((item) => item.name),
  };
}

function bakeryLookupBreadTotalClient(settings, guests) {
  const table = Array.isArray(settings?.breadTable) ? settings.breadTable : [];
  const normalizedGuests = Math.max(0, Number.parseInt(guests, 10) || 0);
  if (!table.length || normalizedGuests <= 0) return 0;
  const found = table.find((item) => item.guests >= normalizedGuests) || table[table.length - 1];
  if (!found) return 0;
  if (normalizeBakeryBaseClient(settings?.selectedBase) === "base-baixa") return Number(found.baseBaixa || 0);
  if (normalizeBakeryBaseClient(settings?.selectedBase) === "base-alta") return Number(found.baseAlta || 0);
  return Number(found.baseMedia || 0);
}

function bakeryAllocateBreadTypesClient(total, types = []) {
  const normalized = (Array.isArray(types) ? types : []).map((item) => ({ ...item, percentage: Number(item.percentage || 0) })).filter((item) => item.name);
  if (!normalized.length) return [];
  const totalBreads = Math.max(0, Number(total || 0));
  const base = normalized.map((item) => {
    const exact = totalBreads * (item.percentage / 100);
    return { ...item, quantity: Math.floor(exact), remainder: exact - Math.floor(exact) };
  });
  let remaining = totalBreads - base.reduce((sum, item) => sum + item.quantity, 0);
  base.sort((a, b) => b.remainder - a.remainder || b.percentage - a.percentage);
  for (let i = 0; i < base.length && remaining > 0; i += 1, remaining -= 1) base[i].quantity += 1;
  return normalized.map((item) => {
    const found = base.find((entry) => entry.name === item.name);
    return { name: item.name, percentage: item.percentage, quantity: found ? found.quantity : 0 };
  });
}

function parseBakeryOptionalCountClient(value) {
  if (value === "" || value === null || value === undefined) return "";
  const num = Number.parseInt(value, 10);
  return Number.isFinite(num) ? Math.max(0, num) : "";
}

function normalizeBakeryDayClient(day = {}, settings = state.bakerySettings) {
  const date = clean(day.date);
  const availableBeds = parseBakeryOptionalCountClient(day.availableBeds ?? day.available_beds);
  const cruzCheckins = parseBakeryOptionalCountClient(day.cruzCheckins ?? day.cruz_checkins);
  const hasAvailableBeds = availableBeds !== "";
  const hasCruzCheckins = cruzCheckins !== "";
  const hostelGuests = hasAvailableBeds ? Math.max(0, Number(settings?.hostelCapacity || 0) - Number(availableBeds)) : "";
  const totalBreads = hostelGuests === "" ? "" : bakeryLookupBreadTotalClient(settings, hostelGuests);
  return {
    date,
    availableBeds,
    cruzCheckins,
    hostelGuests,
    totalBreads,
    breadBreakdown: totalBreads === "" ? bakeryAllocateBreadTypesClient(0, settings?.breadTypes).map((item) => ({ ...item, quantity: "" })) : bakeryAllocateBreadTypesClient(totalBreads, settings?.breadTypes),
    pasteisDeNata: hasCruzCheckins ? cruzCheckins : "",
  };
}

function shouldBlankBakeryFreshDraft(order = {}) {
  const status = clean(order.status).toLowerCase();
  if (status && status !== "open") return false;
  const submittedAt = clean(order.submittedAt || order.submitted_at);
  if (submittedAt) return false;
  const createdAt = clean(order.createdAt || order.created_at);
  const updatedAt = clean(order.updatedAt || order.updated_at);
  return !!createdAt && (!updatedAt || createdAt === updatedAt);
}

function normalizeBakeryOrderClient(order, settings = state.bakerySettings) {
  if (!order) return null;
  const targetDates = Array.isArray(order.targetDates || order.target_dates)
    ? (order.targetDates || order.target_dates).map((item) => clean(item)).filter(Boolean)
    : (Array.isArray(order.days) ? order.days : []).map((item) => clean(item?.date)).filter(Boolean);
  const byDate = new Map((Array.isArray(order.days) ? order.days : []).map((item) => [clean(item.date), item]));
  const blankFreshDefaults = shouldBlankBakeryFreshDraft(order);
  const days = targetDates.map((date) => {
    const rawDay = byDate.get(date) || { date };
    const normalized = normalizeBakeryDayClient(rawDay, settings);
    if (!blankFreshDefaults) return normalized;
    return {
      ...normalized,
      availableBeds: Number(normalized.availableBeds) === 0 ? "" : normalized.availableBeds,
      cruzCheckins: Number(normalized.cruzCheckins) === 0 ? "" : normalized.cruzCheckins,
      hostelGuests: Number(normalized.availableBeds) === 0 ? "" : normalized.hostelGuests,
      totalBreads: Number(normalized.availableBeds) === 0 ? "" : normalized.totalBreads,
      breadBreakdown: Number(normalized.availableBeds) === 0
        ? (Array.isArray(normalized.breadBreakdown) ? normalized.breadBreakdown.map((item) => ({ ...item, quantity: "" })) : [])
        : normalized.breadBreakdown,
      pasteisDeNata: Number(normalized.cruzCheckins) === 0 ? "" : normalized.pasteisDeNata,
    };
  });
  return {
    id: clean(order.id),
    orderNumber: Number(order.orderNumber || order.order_number || 0) || 0,
    status: clean(order.status).toLowerCase() === "submitted" ? "submitted" : "open",
    orderDate: clean(order.orderDate || order.order_date),
    createdAt: clean(order.createdAt || order.created_at),
    updatedAt: clean(order.updatedAt || order.updated_at),
    submittedAt: clean(order.submittedAt || order.submitted_at),
    submittedByName: clean(order.submittedByName || order.submitted_by_name),
    submittedByUserEmail: clean(order.submittedByUserEmail || order.submitted_by_user_email).toLowerCase(),
    targetDates,
    days,
    generatedText: clean(order.generatedText || order.generated_text),
  };
}

function bakeryDateLabel(value) {
  const raw = clean(value);
  if (!raw) return "-";
  try {
    return new Intl.DateTimeFormat("pt-PT", { weekday: "short", day: "2-digit", month: "2-digit", year: "numeric", timeZone: "Europe/Lisbon" }).format(new Date(`${raw}T00:00:00`));
  } catch {
    return raw;
  }
}

function bakeryOrderDatesLabel(order) {
  return (Array.isArray(order?.days) ? order.days : []).map((day) => clean(day.date)).filter(Boolean).join(", ");
}

function bakeryBreadTypeColumnsClient(order = state.bakeryOpenOrder) {
  const seen = new Set();
  const names = [];
  ((state.bakerySettings?.breadTypes) || []).forEach((item) => {
    const name = clean(item?.name);
    if (!name) return;
    const key = name.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    names.push(name);
  });
  (Array.isArray(order?.days) ? order.days : []).forEach((day) => {
    (Array.isArray(day?.breadBreakdown) ? day.breadBreakdown : []).forEach((item) => {
      const name = clean(item?.name);
      if (!name) return;
      const key = name.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      names.push(name);
    });
  });
  return names;
}

function bakeryBreadTypeQuantity(day, breadTypeName) {
  const found = (Array.isArray(day?.breadBreakdown) ? day.breadBreakdown : []).find((item) => clean(item?.name).toLowerCase() === clean(breadTypeName).toLowerCase());
  return found ? found.quantity : "";
}

function buildBakeryGeneratedTextClient(order, name = state.bakerySubmitName) {
  const days = Array.isArray(order?.days) ? order.days : [];
  const lines = [
    `SUBJECT: Lisboa Central Hostel - Encomenda p\u00e3es e bolos para dias ${bakeryOrderDatesLabel(order) || "-"}`,
    "",
    "Bom dia,",
    "",
    "Segue a encomenda de p\u00e3es e bolos:",
    "",
  ];
  days.forEach((day) => {
    lines.push(bakeryDateLabel(day.date));
    (Array.isArray(day.breadBreakdown) ? day.breadBreakdown : []).forEach((item) => lines.push(`${item.name}: ${Number(item.quantity || 0)}`));
    lines.push(`Past\u00e9is de nata: ${Number(day.pasteisDeNata || 0)}`);
    lines.push("");
  });
  lines.push("Cumprimentos,");
  lines.push(clean(name || order?.submittedByName) || "[Name]");
  lines.push("");
  lines.push("Lisboa Central Hostel");
  lines.push("");
  lines.push("+351 309 881 038");
  lines.push("+351 925 222 809");
  lines.push("global@lisboacentralhostel.com");
  lines.push("");
  lines.push("Rua Rodrigues Sampaio 160, 1150-282 Lisboa");
  return lines.join("\n");
}

function buildBakeryGeneratedHtmlClient(order, name = state.bakerySubmitName) {
  const days = Array.isArray(order?.days) ? order.days : [];
  const breadTypes = bakeryBreadTypeColumnsClient(order);
  return `<p><strong>Assunto:</strong> Lisboa Central Hostel - Encomenda p\u00e3es e bolos para dias ${escape(bakeryOrderDatesLabel(order) || "-")}</p>
    <p>Bom dia,</p>
    <p>Segue a encomenda de p\u00e3es e bolos:</p>
    <table>
      <thead>
        <tr>
          <th>Data</th>
          ${breadTypes.map((breadType) => `<th>${escape(breadType)}</th>`).join("")}
          <th style="text-align:center;">Past\u00e9is de nata</th>
        </tr>
      </thead>
      <tbody>${days.map((day) => `<tr>
        <td>${escape(day.date || "-")}</td>
        ${breadTypes.map((breadType) => `<td style="text-align:center;">${escape(String(bakeryBreadTypeQuantity(day, breadType) === "" ? "-" : bakeryBreadTypeQuantity(day, breadType)))}</td>`).join("")}
        <td style="text-align:center;">${escape(String(day.pasteisDeNata === "" ? "-" : day.pasteisDeNata))}</td>
      </tr>`).join("")}</tbody>
    </table>
    <p>Cumprimentos,<br>${escape(clean(name || order?.submittedByName) || "[Name]")}<br><br>
    Lisboa Central Hostel<br><br>
    +351 309 881 038<br>
    +351 925 222 809<br>
    global@lisboacentralhostel.com<br><br>
    Rua Rodrigues Sampaio 160, 1150-282 Lisboa</p>`;
}

function validateBakeryOrderDays(days = [], capacity = Number(state.bakerySettings?.hostelCapacity || 83)) {
  const max = Math.max(0, Number(capacity || 0));
  for (const day of Array.isArray(days) ? days : []) {
    if (day.availableBeds === "" || day.availableBeds === null || day.availableBeds === undefined) {
      return `Available Beds (Hostel) is required for ${bakeryDateLabel(day.date)}.`;
    }
    if (day.cruzCheckins === "" || day.cruzCheckins === null || day.cruzCheckins === undefined) {
      return `Check-ins (Cruz) is required for ${bakeryDateLabel(day.date)}.`;
    }
    if (Number(day.availableBeds) < 0 || Number(day.availableBeds) > max) {
      return `Available Beds (Hostel) must be between 0 and ${max} for ${bakeryDateLabel(day.date)}.`;
    }
    if (Number(day.cruzCheckins) < 0 || Number(day.cruzCheckins) > max) {
      return `Check-ins (Cruz) must be between 0 and ${max} for ${bakeryDateLabel(day.date)}.`;
    }
  }
  return "";
}

function uniqueSortedShoppingValues(items, field) {
  return Array.from(
    new Set(
      (Array.isArray(items) ? items : [])
        .map((item) => clean(item?.[field]))
        .filter(Boolean)
    )
  ).sort((a, b) => a.localeCompare(b));
}

function shoppingHistoryItemKey(item = {}) {
  const id = clean(item.id);
  if (id) return `id:${id}`;
  return `fallback:${clean(item.category).toLowerCase()}::${clean(item.item).toLowerCase()}::${clean(item.supplier).toLowerCase()}`;
}

function dateOnlyIso(value) {
  return clean(value).slice(0, 10);
}

function canReopenShoppingOrderClient(order, historyRows = state.shoppingHistory) {
  const rows = Array.isArray(historyRows) ? historyRows : [];
  const latestId = rows[0]?.id || "";
  const submittedDate = dateOnlyIso(order?.submittedAt || order?.updatedAt || order?.createdAt);
  return !state.shoppingOpenOrder && clean(order?.id) === clean(latestId) && submittedDate === lisbonTodayIsoClient();
}

function isShoppingRecentlyOrdered(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return false;
  const now = new Date();
  return ((now.getTime() - date.getTime()) / (1000 * 60 * 60 * 24)) <= 30;
}

function buildShoppingHistoryStats() {
  const stats = new Map();
  (Array.isArray(state.shoppingHistory) ? state.shoppingHistory : []).forEach((order, orderIndex) => {
    const submittedAt = clean(order.submittedAt || order.updatedAt || order.createdAt);
    shoppingOrderSelectedItems(order).forEach((item) => {
      const key = shoppingHistoryItemKey(item);
      const current = stats.get(key) || {
        count: 0,
        lastOrderedAt: "",
        recent: false,
        frequent: false,
      };
      current.count += 1;
      if (!current.lastOrderedAt) current.lastOrderedAt = submittedAt;
      if (orderIndex < 3 || isShoppingRecentlyOrdered(submittedAt)) current.recent = true;
      stats.set(key, current);
    });
  });
  stats.forEach((value) => {
    value.frequent = value.count >= 3;
  });
  return stats;
}

function getShoppingItemMeta(item, stats = buildShoppingHistoryStats()) {
  return stats.get(shoppingHistoryItemKey(item)) || null;
}

function groupShoppingItemsForDisplay(items, groupBy = state.shoppingFilters.groupBy) {
  const groups = [];
  const map = new Map();
  sortShoppingItemsClient(items, groupBy).forEach((item) => {
    const key = clean(item?.[groupBy]) || (groupBy === "stored" ? "No stored" : "Other");
    if (!map.has(key)) {
      const group = { key, items: [] };
      map.set(key, group);
      groups.push(group);
    }
    map.get(key).items.push(item);
  });
  return groups;
}

function collectShoppingHistoryItems(rows = []) {
  return (Array.isArray(rows) ? rows : []).reduce((acc, order) => {
    shoppingOrderSelectedItems(order).forEach((item) => acc.push(item));
    return acc;
  }, []);
}

function renderShoppingCurrentFilters(order) {
  const items = Array.isArray(order?.items) ? order.items : [];
  const categories = uniqueSortedShoppingValues(items, "category");
  const storedValues = uniqueSortedShoppingValues(items, "stored");
  if (state.shoppingFilters.category && !categories.includes(state.shoppingFilters.category)) state.shoppingFilters.category = "";
  if (state.shoppingFilters.stored && !storedValues.includes(state.shoppingFilters.stored)) state.shoppingFilters.stored = "";
  if (!["category", "stored"].includes(clean(state.shoppingFilters.groupBy))) state.shoppingFilters.groupBy = "category";
  if (els.shoppingFilterCategory) {
    els.shoppingFilterCategory.innerHTML = [`<option value="">All categories</option>`, ...categories.map((value) => `<option value="${escape(value)}">${escape(value)}</option>`)].join("");
    els.shoppingFilterCategory.value = state.shoppingFilters.category;
  }
  if (els.shoppingFilterStored) {
    els.shoppingFilterStored.innerHTML = [`<option value="">All stored</option>`, ...storedValues.map((value) => `<option value="${escape(value)}">${escape(value)}</option>`)].join("");
    els.shoppingFilterStored.value = state.shoppingFilters.stored;
  }
  if (els.shoppingGroupBy) els.shoppingGroupBy.value = state.shoppingFilters.groupBy;
}

function getFilteredShoppingItems(order) {
  const items = Array.isArray(order?.items) ? order.items : [];
  const category = clean(state.shoppingFilters.category);
  const stored = clean(state.shoppingFilters.stored);
  return sortShoppingItemsClient(items.filter((item) => {
    if (category && clean(item.category) !== category) return false;
    if (stored && clean(item.stored) !== stored) return false;
    return true;
  }), state.shoppingFilters.groupBy);
}

function setShoppingCurrentStatus(text) {
  if (els.shoppingCurrentStatus) els.shoppingCurrentStatus.textContent = text;
}

function setShoppingSubmitStatus(text) {
  if (els.shoppingSubmitStatus) els.shoppingSubmitStatus.textContent = text;
}

function setShoppingHistoryStatus(text) {
  if (els.shoppingHistoryStatus) els.shoppingHistoryStatus.textContent = text;
}

function setShoppingSettingsStatus(text) {
  if (els.shoppingSettingsStatus) els.shoppingSettingsStatus.textContent = text;
}

function setShoppingDetailStatus(text) {
  if (els.shoppingDetailStatus) els.shoppingDetailStatus.textContent = text;
}

async function loadShoppingSettings({ silent = false } = {}) {
  try {
    const result = await api("/api/shopping-settings");
    state.shoppingSettings = normalizeShoppingSettingsClient(result.settings);
    renderShoppingSettings();
    if (!silent) setShoppingSettingsStatus(`Loaded ${state.shoppingSettings.items.length} shopping items.`);
  } catch (e) {
    state.shoppingSettings = clone(DEFAULT_SHOPPING_SETTINGS);
    renderShoppingSettings();
    if (!silent) setShoppingSettingsStatus(`Using default shopping settings (${e.message}).`);
  }
}

async function loadShoppingData({ silent = false } = {}) {
  try {
    const result = await api("/api/shopping");
    if (result.settings) {
      state.shoppingSettings = normalizeShoppingSettingsClient(result.settings);
      state.shoppingSettingsLoaded = true;
    }
    state.shoppingOpenOrder = normalizeShoppingOrderClient(result.openOrder, state.shoppingSettings.items);
    state.shoppingHistory = (Array.isArray(result.history) ? result.history : [])
      .map((order) => normalizeShoppingOrderClient(order, state.shoppingSettings.items))
      .filter(Boolean);
    if (!state.shoppingOpenOrder) {
      state.shoppingSubmitName = "";
      state.shoppingSubmitNotes = "";
      state.shoppingSubmitPromptOpen = false;
    }
    let renderError = "";
    try {
      renderShopping();
    } catch (renderIssue) {
      renderError = renderIssue?.message || "Could not render Shopping app.";
    }
    try {
      renderShoppingSettings();
    } catch (settingsIssue) {
      renderError = renderError || settingsIssue?.message || "Could not render Shopping settings.";
    }
    if (!silent) {
      if (renderError) setShoppingCurrentStatus(`Shopping loaded with display issue: ${renderError}`);
      else setShoppingCurrentStatus(state.shoppingOpenOrder ? "Open shopping order loaded." : "No open shopping order.");
    }
  } catch (e) {
    try {
      renderShoppingSettings();
    } catch {}
    if (!silent) setShoppingCurrentStatus(`Could not load shopping data: ${e.message}`);
  }
}

function setShoppingTab(tab) {
  state.shoppingTab = tab === "history" ? "history" : "current";
  if (els.shoppingTabCurrent) {
    els.shoppingTabCurrent.classList.toggle("active-tab", state.shoppingTab === "current");
    els.shoppingTabCurrent.classList.toggle("ghost", state.shoppingTab !== "current");
  }
  if (els.shoppingTabHistory) {
    els.shoppingTabHistory.classList.toggle("active-tab", state.shoppingTab === "history");
    els.shoppingTabHistory.classList.toggle("ghost", state.shoppingTab !== "history");
  }
  if (els.shoppingPanelCurrent) els.shoppingPanelCurrent.hidden = state.shoppingTab !== "current";
  if (els.shoppingPanelHistory) els.shoppingPanelHistory.hidden = state.shoppingTab !== "history";
}

function shouldShowShoppingAlert() {
  if (!canApp("shopping")) return false;
  const weekdays = normalizeShoppingWeekdaysClient(state.shoppingSettings?.mandatoryWeekdays);
  if (!weekdays.length) return false;
  const today = new Date();
  const weekdayKey = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"][today.getDay()];
  if (!weekdays.includes(weekdayKey)) return false;
  const todayIso = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(today);
  return !state.shoppingHistory.some((order) => clean(order.submittedAt).slice(0, 10) === todayIso);
}

function lisbonTodayIsoClient() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
}

function easterSundayClient(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(Date.UTC(year, month - 1, day));
}

function shiftUtcDaysClient(date, days) {
  const clone = new Date(date.getTime());
  clone.setUTCDate(clone.getUTCDate() + days);
  return clone;
}

function lisbonHolidaySetClient(year) {
  const easter = easterSundayClient(year);
  const goodFriday = shiftUtcDaysClient(easter, -2);
  const corpusChristi = shiftUtcDaysClient(easter, 60);
  return new Set([
    `${year}-01-01`,
    `${year}-04-25`,
    `${year}-05-01`,
    `${year}-06-10`,
    `${year}-06-13`,
    `${year}-08-15`,
    `${year}-10-05`,
    `${year}-11-01`,
    `${year}-12-01`,
    `${year}-12-08`,
    `${year}-12-25`,
    goodFriday.toISOString().slice(0, 10),
    corpusChristi.toISOString().slice(0, 10),
  ]);
}

function isWorkingDayLisbonClient(iso) {
  const raw = clean(iso);
  if (!raw) return false;
  const date = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(date.getTime())) return false;
  const weekday = date.getDay();
  if (weekday === 0 || weekday === 6) return false;
  return !lisbonHolidaySetClient(date.getFullYear()).has(raw);
}

function shouldShowBakeryAlert() {
  if (!canApp("bakery")) return false;
  const todayIso = lisbonTodayIsoClient();
  if (!isWorkingDayLisbonClient(todayIso)) return false;
  return !state.bakeryHistory.some((order) => clean(order.submittedAt).slice(0, 10) === todayIso);
}

function shouldShowLaundryAlert() {
  if (!canApp("laundry")) return false;
  return laundryHasMissingSentRecords() || laundryHasOverduePendingReceipts();
}

function hexToShoppingRowColor(hex, alpha = 0.5) {
  const raw = clean(hex).replace("#", "");
  if (!/^[0-9a-f]{6}$/i.test(raw)) return "";
  const r = Number.parseInt(raw.slice(0, 2), 16);
  const g = Number.parseInt(raw.slice(2, 4), 16);
  const b = Number.parseInt(raw.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function getShoppingCategoryColor(category) {
  const normalized = normalizeShoppingCategoryClient(category);
  const configured = state.shoppingSettings?.categoryColors?.[normalized];
  return normalizeShoppingColorClient(configured, DEFAULT_SHOPPING_CATEGORY_COLORS[normalized] || "#F3E7DB");
}

function renderShoppingCurrentRowsLegacy(order) {
  const items = getFilteredShoppingItems(order);
  if (els.shoppingOpenRows) els.shoppingOpenRows.innerHTML = "";
  if (els.shoppingMobileCards) els.shoppingMobileCards.innerHTML = "";
  if (!items.length) {
    if (els.shoppingOpenRows) els.shoppingOpenRows.innerHTML = '<tr><td colspan="6" class="empty">No shopping items match the current filters.</td></tr>';
    if (els.shoppingMobileCards) els.shoppingMobileCards.innerHTML = '<div class="services-mobile-empty">No shopping items match the current filters.</div>';
    return;
  }
  items.forEach((item) => {
    const rowColor = hexToShoppingRowColor(getShoppingCategoryColor(item.category), 0.15);
    const quantityPlaceholder = item.order && item.quantityRequired ? "Required" : "";
    const quantityDisabled = item.order ? "" : "disabled";
    const tr = document.createElement("tr");
    if (rowColor) tr.style.background = rowColor;
    tr.innerHTML = `<td>${escape(item.category || "-")}</td>
      <td>${escape(item.item || "-")}</td>
      <td>${escape(item.supplier || "-")}</td>
      <td>${escape(item.stored || "-")}</td>
      <td><input class="shopping-existing-qty-input${item.order ? "" : " is-disabled"}" data-shopping-item-id="${escape(item.id)}" data-shopping-field="existingQuantity" type="text" value="${escape(item.existingQuantity || "")}" placeholder="${escape(quantityPlaceholder)}" ${quantityDisabled} /></td>
      <td><label class="status-toggle"><input data-shopping-item-id="${escape(item.id)}" data-shopping-field="order" type="checkbox" ${item.order ? "checked" : ""} /><span>Order</span></label></td>`;
    els.shoppingOpenRows.appendChild(tr);
    if (els.shoppingMobileCards) {
      const card = document.createElement("article");
      card.className = `shopping-mobile-card${item.order ? " selected-card" : ""}`;
      if (rowColor) card.style.background = rowColor;
      card.innerHTML = `<div class="shopping-mobile-row">
          <div class="shopping-mobile-main">
            <div class="service-mobile-request">${escape(item.item || "-")}</div>
            <div class="service-mobile-type">${escape(item.category || "-")}${item.supplier ? ` · ${escape(item.supplier)}` : ""}</div>
          </div>
          <label class="shopping-mobile-inline-order" aria-label="Order item">
            <input data-shopping-item-id="${escape(item.id)}" data-shopping-field="order" type="checkbox" ${item.order ? "checked" : ""} />
          </label>
          <input class="shopping-existing-qty-input shopping-mobile-inline-qty${item.order ? "" : " is-disabled"}" data-shopping-item-id="${escape(item.id)}" data-shopping-field="existingQuantity" type="text" value="${escape(item.existingQuantity || "")}" placeholder="${escape(quantityPlaceholder)}" ${quantityDisabled} />
        </div>`;
      els.shoppingMobileCards.appendChild(card);
    }
  });
}

function renderShoppingHistoryRowsLegacy() {
  const rows = Array.isArray(state.shoppingHistory) ? state.shoppingHistory : [];
  if (els.shoppingHistoryRows) els.shoppingHistoryRows.innerHTML = "";
  if (els.shoppingHistoryMobileCards) els.shoppingHistoryMobileCards.innerHTML = "";
  els.shoppingHistoryCount.textContent = `${rows.length} order${rows.length === 1 ? "" : "s"}`;
  if (!rows.length) {
    els.shoppingHistoryRows.innerHTML = '<tr><td colspan="4" class="empty">No submitted shopping orders yet.</td></tr>';
    if (els.shoppingHistoryMobileCards) {
      els.shoppingHistoryMobileCards.innerHTML = '<div class="services-mobile-empty">No submitted shopping orders yet.</div>';
    }
    return;
  }
  rows.forEach((order, index) => {
    const canReopen = !state.shoppingOpenOrder && index === 0;
    const tr = document.createElement("tr");
    tr.className = "clickable-row";
    tr.dataset.shoppingHistoryId = order.id;
    tr.innerHTML = `<td>${escape(formatDateTimeShort(order.submittedAt || order.updatedAt || order.createdAt))}</td>
      <td>${escape(order.submittedByName || "-")}</td>
      <td>${escape(String(order.orderedCount || 0))}</td>
      <td>${canReopen ? `<button type="button" class="ghost" data-action="reopen-shopping-order" data-id="${escape(order.id)}">Reopen</button>` : ""}</td>`;
    els.shoppingHistoryRows.appendChild(tr);
    if (els.shoppingHistoryMobileCards) {
      const card = document.createElement("article");
      card.className = "shopping-history-card";
      card.dataset.shoppingHistoryId = order.id;
      card.innerHTML = `<div class="service-mobile-header">
          <div>
            <div class="service-mobile-request">${escape(order.submittedByName || "-")}</div>
            <div class="service-mobile-type">${escape(formatDateTimeShort(order.submittedAt || order.updatedAt || order.createdAt))}</div>
          </div>
          <div class="review-mobile-score">${escape(String(order.orderedCount || 0))}</div>
        </div>
        ${canReopen ? `<div class="row-actions"><button type="button" class="ghost" data-action="reopen-shopping-order" data-id="${escape(order.id)}">Reopen</button></div>` : ""}`;
      els.shoppingHistoryMobileCards.appendChild(card);
    }
  });
}

function getShoppingDetailOrder() {
  return state.shoppingHistory.find((item) => clean(item.id) === clean(state.shoppingSelectedHistoryId)) || null;
}

function shoppingOrderExportFileStem(order) {
  const number = String(order?.orderNumber || "order").padStart(4, "0");
  return `shopping_order_${number}`;
}

function shoppingOrderMetaRows(order) {
  return [
    ["Order #", String(order?.orderNumber || "-")],
    ["Order Date", formatDateTimeShort(order?.submittedAt || order?.updatedAt || order?.createdAt)],
    ["Name", order?.submittedByName || "-"],
    ["Number Items", String(order?.orderedCount || 0)],
    ["Notes", order?.notes || "-"],
  ];
}

function shoppingOrderSelectedItems(order) {
  return sortShoppingItemsClient((Array.isArray(order?.items) ? order.items : []).filter((item) => item.order), "category");
}

function compareShoppingItemsClient(a, b, groupBy = "category") {
  const primaryField = groupBy === "stored" ? "stored" : "category";
  const secondaryField = primaryField === "stored" ? "category" : "stored";
  const primaryCompare = clean(a?.[primaryField]).localeCompare(clean(b?.[primaryField]), undefined, { sensitivity: "base" });
  if (primaryCompare !== 0) return primaryCompare;
  const secondaryCompare = clean(a?.[secondaryField]).localeCompare(clean(b?.[secondaryField]), undefined, { sensitivity: "base" });
  if (secondaryCompare !== 0) return secondaryCompare;
  const itemCompare = clean(a?.item).localeCompare(clean(b?.item), undefined, { sensitivity: "base" });
  if (itemCompare !== 0) return itemCompare;
  return clean(a?.supplier).localeCompare(clean(b?.supplier), undefined, { sensitivity: "base" });
}

function sortShoppingItemsClient(items, groupBy = state.shoppingFilters.groupBy) {
  return [...(Array.isArray(items) ? items : [])].sort((a, b) => compareShoppingItemsClient(a, b, groupBy));
}

function renderShoppingCurrentRows(order) {
  const items = getFilteredShoppingItems(order);
  const stats = buildShoppingHistoryStats();
  const groupBy = state.shoppingFilters.groupBy === "stored" ? "stored" : "category";
  const groups = groupShoppingItemsForDisplay(items, groupBy);
  const groupLabel = groupBy === "stored" ? "Stored" : "Category";
  if (els.shoppingOpenRows) els.shoppingOpenRows.innerHTML = "";
  if (els.shoppingMobileCards) els.shoppingMobileCards.innerHTML = "";
  if (!items.length) {
    if (els.shoppingOpenRows) els.shoppingOpenRows.innerHTML = '<tr><td colspan="6" class="empty">No shopping items match the current filters.</td></tr>';
    if (els.shoppingMobileCards) els.shoppingMobileCards.innerHTML = '<div class="services-mobile-empty">No shopping items match the current filters.</div>';
    return;
  }
  groups.forEach((group) => {
    if (els.shoppingOpenRows) {
      const headerRow = document.createElement("tr");
      headerRow.className = "shopping-group-row";
      headerRow.innerHTML = `<td colspan="6">${escape(groupLabel)}: ${escape(group.key || "-")}</td>`;
      els.shoppingOpenRows.appendChild(headerRow);
    }
    if (els.shoppingMobileCards) {
      const header = document.createElement("div");
      header.className = "shopping-group-label";
      header.textContent = `${groupLabel}: ${group.key || "-"}`;
      els.shoppingMobileCards.appendChild(header);
    }
    group.items.forEach((item) => {
      const rowColor = hexToShoppingRowColor(getShoppingCategoryColor(item.category), 0.15);
      const meta = getShoppingItemMeta(item, stats);
      const indicators = [
        meta?.recent ? '<span class="shopping-item-indicator">Recent</span>' : "",
        meta?.frequent ? '<span class="shopping-item-indicator">Frequent</span>' : "",
      ].filter(Boolean).join("");
      const lastOrdered = meta?.lastOrderedAt
        ? `<div class="shopping-item-meta">Last ordered: ${escape(formatDateOnly(dateOnlyIso(meta.lastOrderedAt)))} ${indicators}</div>`
        : (indicators ? `<div class="shopping-item-meta">${indicators}</div>` : "");
      const quantityPlaceholder = item.order && item.quantityRequired ? "Required" : "";
      const quantityDisabled = item.order ? "" : "disabled";
      const tr = document.createElement("tr");
      if (rowColor) tr.style.background = rowColor;
      tr.innerHTML = `<td>${escape(item.category || "-")}</td>
        <td><div class="shopping-item-name">${escape(item.item || "-")}</div>${lastOrdered}</td>
        <td>${escape(item.supplier || "-")}</td>
        <td>${escape(item.stored || "-")}</td>
        <td><input class="shopping-existing-qty-input${item.order ? "" : " is-disabled"}" data-shopping-item-id="${escape(item.id)}" data-shopping-field="existingQuantity" type="text" value="${escape(item.existingQuantity || "")}" placeholder="${escape(quantityPlaceholder)}" ${quantityDisabled} /></td>
        <td><label class="status-toggle"><input data-shopping-item-id="${escape(item.id)}" data-shopping-field="order" type="checkbox" ${item.order ? "checked" : ""} /><span>Order</span></label></td>`;
      els.shoppingOpenRows.appendChild(tr);
      if (els.shoppingMobileCards) {
        const card = document.createElement("article");
        card.className = `shopping-mobile-card${item.order ? " selected-card" : ""}`;
        if (rowColor) card.style.background = rowColor;
        card.innerHTML = `<div class="shopping-mobile-row">
            <div class="shopping-mobile-main">
              <div class="service-mobile-request">${escape(item.item || "-")}</div>
              <div class="service-mobile-type">${escape(item.category || "-")}${item.supplier ? ` Â· ${escape(item.supplier)}` : ""}</div>
              ${lastOrdered}
            </div>
            <label class="shopping-mobile-inline-order" aria-label="Order item">
              <input data-shopping-item-id="${escape(item.id)}" data-shopping-field="order" type="checkbox" ${item.order ? "checked" : ""} />
            </label>
            <input class="shopping-existing-qty-input shopping-mobile-inline-qty${item.order ? "" : " is-disabled"}" data-shopping-item-id="${escape(item.id)}" data-shopping-field="existingQuantity" type="text" value="${escape(item.existingQuantity || "")}" placeholder="${escape(quantityPlaceholder)}" ${quantityDisabled} />
          </div>`;
        els.shoppingMobileCards.appendChild(card);
      }
    });
  });
}

function renderShoppingHistoryFilters() {
  const rows = Array.isArray(state.shoppingHistory) ? state.shoppingHistory : [];
  const historyItems = collectShoppingHistoryItems(rows);
  const categories = uniqueSortedShoppingValues(historyItems, "category");
  const suppliers = uniqueSortedShoppingValues(historyItems, "supplier");
  if (state.shoppingHistoryFilters.category && !categories.includes(state.shoppingHistoryFilters.category)) state.shoppingHistoryFilters.category = "";
  if (state.shoppingHistoryFilters.supplier && !suppliers.includes(state.shoppingHistoryFilters.supplier)) state.shoppingHistoryFilters.supplier = "";
  if (els.shoppingHistoryDateFrom) els.shoppingHistoryDateFrom.value = state.shoppingHistoryFilters.dateFrom;
  if (els.shoppingHistoryDateTo) els.shoppingHistoryDateTo.value = state.shoppingHistoryFilters.dateTo;
  if (els.shoppingHistoryName) els.shoppingHistoryName.value = state.shoppingHistoryFilters.name;
  if (els.shoppingHistoryCategory) {
    els.shoppingHistoryCategory.innerHTML = [`<option value="">All categories</option>`, ...categories.map((value) => `<option value="${escape(value)}">${escape(value)}</option>`)].join("");
    els.shoppingHistoryCategory.value = state.shoppingHistoryFilters.category;
  }
  if (els.shoppingHistorySupplier) {
    els.shoppingHistorySupplier.innerHTML = [`<option value="">All suppliers</option>`, ...suppliers.map((value) => `<option value="${escape(value)}">${escape(value)}</option>`)].join("");
    els.shoppingHistorySupplier.value = state.shoppingHistoryFilters.supplier;
  }
}

function getFilteredShoppingHistoryRows() {
  const rows = Array.isArray(state.shoppingHistory) ? state.shoppingHistory : [];
  const dateFrom = clean(state.shoppingHistoryFilters.dateFrom);
  const dateTo = clean(state.shoppingHistoryFilters.dateTo);
  const name = clean(state.shoppingHistoryFilters.name).toLowerCase();
  const category = clean(state.shoppingHistoryFilters.category);
  const supplier = clean(state.shoppingHistoryFilters.supplier);
  return rows.filter((order) => {
    const submittedDate = dateOnlyIso(order.submittedAt || order.updatedAt || order.createdAt);
    if (dateFrom && submittedDate && submittedDate < dateFrom) return false;
    if (dateTo && submittedDate && submittedDate > dateTo) return false;
    if (name && !clean(order.submittedByName).toLowerCase().includes(name)) return false;
    const selectedItems = shoppingOrderSelectedItems(order);
    if (category && !selectedItems.some((item) => clean(item.category) === category)) return false;
    if (supplier && !selectedItems.some((item) => clean(item.supplier) === supplier)) return false;
    return true;
  });
}

function renderShoppingHistoryRows() {
  const totalRows = Array.isArray(state.shoppingHistory) ? state.shoppingHistory : [];
  const rows = getFilteredShoppingHistoryRows();
  if (els.shoppingHistoryRows) els.shoppingHistoryRows.innerHTML = "";
  if (els.shoppingHistoryMobileCards) els.shoppingHistoryMobileCards.innerHTML = "";
  renderShoppingHistoryFilters();
  if (!totalRows.length) {
    els.shoppingHistoryCount.textContent = "0 orders";
    els.shoppingHistoryRows.innerHTML = '<tr><td colspan="4" class="empty">No submitted shopping orders yet.</td></tr>';
    if (els.shoppingHistoryMobileCards) {
      els.shoppingHistoryMobileCards.innerHTML = '<div class="services-mobile-empty">No submitted shopping orders yet.</div>';
    }
    return;
  }
  els.shoppingHistoryCount.textContent =
    rows.length === totalRows.length ? `${rows.length} order${rows.length === 1 ? "" : "s"}` : `${rows.length} of ${totalRows.length} orders`;
  if (!rows.length) {
    els.shoppingHistoryRows.innerHTML = '<tr><td colspan="4" class="empty">No shopping orders match the current filters.</td></tr>';
    if (els.shoppingHistoryMobileCards) {
      els.shoppingHistoryMobileCards.innerHTML = '<div class="services-mobile-empty">No shopping orders match the current filters.</div>';
    }
    return;
  }
  const latestId = totalRows[0]?.id || "";
  rows.forEach((order) => {
    const canReopen = canReopenShoppingOrderClient(order, totalRows);
    const canCopy = !state.shoppingOpenOrder;
    const tr = document.createElement("tr");
    tr.className = "clickable-row";
    tr.dataset.shoppingHistoryId = order.id;
    tr.innerHTML = `<td>${escape(formatDateTimeShort(order.submittedAt || order.updatedAt || order.createdAt))}</td>
      <td>${escape(order.submittedByName || "-")}</td>
      <td>${escape(String(order.orderedCount || 0))}</td>
      <td>${canCopy ? `<button type="button" class="ghost" data-action="copy-shopping-order" data-id="${escape(order.id)}">Copy as Draft</button>` : ""}${canReopen ? ` <button type="button" class="ghost" data-action="reopen-shopping-order" data-id="${escape(order.id)}">Reopen</button>` : ""}</td>`;
    els.shoppingHistoryRows.appendChild(tr);
    if (els.shoppingHistoryMobileCards) {
      const card = document.createElement("article");
      card.className = "shopping-history-card";
      card.dataset.shoppingHistoryId = order.id;
      card.innerHTML = `<div class="service-mobile-header">
          <div>
            <div class="service-mobile-request">${escape(order.submittedByName || "-")}</div>
            <div class="service-mobile-type">${escape(formatDateTimeShort(order.submittedAt || order.updatedAt || order.createdAt))}</div>
          </div>
          <div class="review-mobile-score">${escape(String(order.orderedCount || 0))}</div>
        </div>
        ${(canCopy || canReopen) ? `<div class="row-actions">${canCopy ? `<button type="button" class="ghost" data-action="copy-shopping-order" data-id="${escape(order.id)}">Copy as Draft</button>` : ""}${canReopen ? ` <button type="button" class="ghost" data-action="reopen-shopping-order" data-id="${escape(order.id)}">Reopen</button>` : ""}</div>` : ""}`;
      els.shoppingHistoryMobileCards.appendChild(card);
    }
  });
}

function blendShoppingColorOnWhiteClient(hex, alpha = 0.15) {
  const normalized = normalizeShoppingColorClient(hex, "#F3E7DB");
  const value = normalized.replace("#", "");
  const r = Number.parseInt(value.slice(0, 2), 16);
  const g = Number.parseInt(value.slice(2, 4), 16);
  const b = Number.parseInt(value.slice(4, 6), 16);
  const blend = (channel) => Math.round((255 * (1 - alpha)) + (channel * alpha));
  return `#${[blend(r), blend(g), blend(b)].map((item) => item.toString(16).padStart(2, "0")).join("")}`;
}

function buildShoppingOrderExcelHtmlClient(order) {
  const rows = shoppingOrderSelectedItems(order);
  const metaRows = shoppingOrderMetaRows(order);
  return `<!doctype html><html><head><meta charset="utf-8"><style>
    @page { size: A4 portrait; margin: 12mm; }
    body { font-family: Calibri, Arial, sans-serif; color: #222; font-size: 11px; }
    h1 { font-size: 18px; margin: 0 0 8px; color: #3b2f24; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; mso-page-orientation: portrait; }
    th, td { border: 1px solid #cfc7bf; padding: 6px 7px; vertical-align: top; text-align: left; }
    th { background: #e8ded4; font-weight: 700; }
    .meta td:first-child { width: 150px; font-weight: 700; background: #faf7f2; }
    .items th:nth-child(1) { width: 16%; }
    .items th:nth-child(2) { width: 28%; }
    .items th:nth-child(3) { width: 20%; }
    .items th:nth-child(4) { width: 16%; }
    .items th:nth-child(5) { width: 12%; }
    .items th:nth-child(6) { width: 8%; }
  </style></head><body>
    <h1>Shopping Order Detail</h1>
    <table class="meta"><tbody>${metaRows.map(([label, value]) => `<tr><td>${escape(label)}</td><td>${escape(value)}</td></tr>`).join("")}</tbody></table>
    <table class="items"><thead><tr><th>Category</th><th>Item</th><th>Supplier</th><th>Stored</th><th>Existing Quantity</th><th>Order</th></tr></thead>
    <tbody>${rows.map((item) => `<tr style="background:${escape(blendShoppingColorOnWhiteClient(getShoppingCategoryColor(item.category), 0.15))}"><td>${escape(item.category || "-")}</td><td>${escape(item.item || "-")}</td><td>${escape(item.supplier || "-")}</td><td>${escape(item.stored || "-")}</td><td>${escape(item.existingQuantity || "-")}</td><td>Yes</td></tr>`).join("")}</tbody></table>
  </body></html>`;
}

function wrapShoppingPdfTextClient(text, width = 88) {
  const words = String(text ?? "").split(/\s+/).filter(Boolean);
  if (!words.length) return [""];
  const lines = [];
  let current = "";
  words.forEach((word) => {
    const candidate = current ? `${current} ${word}` : word;
    if (candidate.length > width && current) {
      lines.push(current);
      current = word;
    } else {
      current = candidate;
    }
  });
  if (current) lines.push(current);
  return lines;
}

function pdfEscapeClient(text) {
  return String(text ?? "").replaceAll("\\", "\\\\").replaceAll("(", "\\(").replaceAll(")", "\\)");
}

function normalizePdfTextClient(value) {
  return String(value ?? "")
    .replaceAll("\u00a0", " ")
    .replaceAll("\u2018", "'")
    .replaceAll("\u2019", "'")
    .replaceAll("\u201c", '"')
    .replaceAll("\u201d", '"')
    .replaceAll("\u2013", "-")
    .replaceAll("\u2014", "-")
    .replaceAll("\u2026", "...")
    .replaceAll("\u2022", "-")
    .replaceAll("\u200b", "")
    .split("")
    .map((char) => (char.charCodeAt(0) <= 255 ? char : "?"))
    .join("");
}

function pdfRgbClient(hex) {
  const normalized = normalizeShoppingColorClient(hex, "#F3E7DB").replace("#", "");
  return [
    Number.parseInt(normalized.slice(0, 2), 16) / 255,
    Number.parseInt(normalized.slice(2, 4), 16) / 255,
    Number.parseInt(normalized.slice(4, 6), 16) / 255,
  ];
}

function drawPdfRectClient(commands, x, y, width, height, fillHex, strokeHex = "#cfc7bf") {
  const [fr, fg, fb] = pdfRgbClient(fillHex);
  const [sr, sg, sb] = pdfRgbClient(strokeHex);
  commands.push(`${fr.toFixed(3)} ${fg.toFixed(3)} ${fb.toFixed(3)} rg`);
  commands.push(`${x.toFixed(2)} ${y.toFixed(2)} ${width.toFixed(2)} ${height.toFixed(2)} re f`);
  commands.push(`${sr.toFixed(3)} ${sg.toFixed(3)} ${sb.toFixed(3)} RG`);
  commands.push(`0.6 w ${x.toFixed(2)} ${y.toFixed(2)} ${width.toFixed(2)} ${height.toFixed(2)} re S`);
}

function buildShoppingOrderPdfBytesClient(order) {
  const rows = shoppingOrderSelectedItems(order);
  const pageWidth = 595;
  const pageHeight = 842;
  const margin = 32;
  const colWidths = [78, 160, 105, 78, 78, 52];
  const headers = ["Category", "Item", "Supplier", "Stored", "Existing Qty", "Order"];
  const metaRows = shoppingOrderMetaRows(order);
  const pages = [];
  let commands = [];
  let y = pageHeight - margin;
  const startPage = () => {
    commands = [];
    y = pageHeight - margin;
    commands.push("0 0 0 rg");
    commands.push("BT");
    commands.push(`/F2 14 Tf 1 0 0 1 ${margin} ${y} Tm (${pdfEscapeClient(normalizePdfTextClient("Shopping Order Detail"))}) Tj`);
    commands.push("ET");
    y -= 20;
    metaRows.forEach(([label, value]) => {
      commands.push("0 0 0 rg");
      commands.push("BT");
      commands.push(`/F2 8 Tf 1 0 0 1 ${margin} ${y} Tm (${pdfEscapeClient(normalizePdfTextClient(`${label}:`))}) Tj`);
      commands.push(`/F1 8 Tf 1 0 0 1 ${margin + 90} ${y} Tm (${pdfEscapeClient(normalizePdfTextClient(value))}) Tj`);
      commands.push("ET");
      y -= 11;
    });
    y -= 4;
  };
  const flushPage = () => {
    pages.push(commands.join("\n"));
  };
  const drawHeader = () => {
    const rowHeight = 15;
    let x = margin;
    headers.forEach((header, index) => {
      drawPdfRectClient(commands, x, y - rowHeight, colWidths[index], rowHeight, "#f1ece6");
      commands.push("0 0 0 rg");
      commands.push("BT");
      commands.push(`/F2 7.5 Tf 1 0 0 1 ${x + 3} ${y - 10.5} Tm (${pdfEscapeClient(normalizePdfTextClient(header))}) Tj`);
      commands.push("ET");
      x += colWidths[index];
    });
    y -= rowHeight;
  };
  const wrapCell = (text, width) => wrapShoppingPdfTextClient(text, Math.max(8, Math.floor((width - 6) / 5.0)));
  startPage();
  drawHeader();
  rows.forEach((item) => {
    const cells = [
      item.category || "-",
      item.item || "-",
      item.supplier || "-",
      item.stored || "-",
      item.existingQuantity || "-",
      "Yes",
    ];
    const wrapped = cells.map((cell, index) => wrapCell(cell, colWidths[index]));
    const rowLines = Math.max(...wrapped.map((cell) => cell.length), 1);
    const rowHeight = Math.max(14, rowLines * 8.5 + 4);
    if (y - rowHeight < margin) {
      flushPage();
      startPage();
      drawHeader();
    }
    let x = margin;
    const fill = blendShoppingColorOnWhiteClient(getShoppingCategoryColor(item.category), 0.10);
    wrapped.forEach((cellLines, index) => {
      drawPdfRectClient(commands, x, y - rowHeight, colWidths[index], rowHeight, fill);
      cellLines.forEach((line, lineIndex) => {
        commands.push("0 0 0 rg");
        commands.push("BT");
        commands.push(`/F1 7.5 Tf 1 0 0 1 ${x + 3} ${y - 10.5 - lineIndex * 8.5} Tm (${pdfEscapeClient(normalizePdfTextClient(line))}) Tj`);
        commands.push("ET");
      });
      x += colWidths[index];
    });
    y -= rowHeight;
  });
  if (!rows.length) {
    const rowHeight = 18;
    drawPdfRectClient(commands, margin, y - rowHeight, colWidths.reduce((sum, value) => sum + value, 0), rowHeight, "#faf7f2");
    commands.push("0 0 0 rg");
    commands.push("BT");
    commands.push(`/F1 8 Tf 1 0 0 1 ${margin + 3} ${y - 11.5} Tm (${pdfEscapeClient(normalizePdfTextClient("No selected items in this order."))}) Tj`);
    commands.push("ET");
    y -= rowHeight;
  }
  flushPage();
  const objects = [];
  objects.push("<< /Type /Catalog /Pages 2 0 R >>");
  objects.push(`<< /Type /Pages /Count ${pages.length} /Kids [${pages.map((_, i) => `${3 + i * 2} 0 R`).join(" ")}] >>`);
  pages.forEach((stream, index) => {
    const pageObj = 3 + index * 2;
    const contentObj = pageObj + 1;
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${3 + pages.length * 2} 0 R /F2 ${4 + pages.length * 2} 0 R >> >> /Contents ${contentObj} 0 R >>`);
    objects.push(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
  });
  objects.push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  objects.push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>");
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((obj, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${obj}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach((offset) => {
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return new Uint8Array(Array.from(pdf, (char) => char.charCodeAt(0) & 0xff));
}

function exportShoppingDetailToExcel() {
  const order = getShoppingDetailOrder();
  if (!order) {
    setShoppingDetailStatus("Could not find the selected shopping order.");
    return;
  }
  downloadBlob(`${shoppingOrderExportFileStem(order)}.xls`, buildShoppingOrderExcelHtmlClient(order), "application/vnd.ms-excel;charset=utf-8;");
  showToast("Shopping order exported to Excel.", "success");
}

function exportShoppingDetailToPdf() {
  const order = getShoppingDetailOrder();
  if (!order) {
    setShoppingDetailStatus("Could not find the selected shopping order.");
    return;
  }
  downloadBlob(`${shoppingOrderExportFileStem(order)}.pdf`, buildShoppingOrderPdfBytesClient(order), "application/pdf");
  showToast("Shopping order exported to PDF.", "success");
}

function renderShoppingDetail(order) {
  if (!order) {
    els.shoppingDetailBody.className = "review-detail empty";
    els.shoppingDetailBody.textContent = "Select an order to see the detail.";
    return;
  }
  const selectedItems = shoppingOrderSelectedItems(order);
  const meta = [
    ["Order Date", formatDateTimeShort(order.submittedAt || order.updatedAt || order.createdAt)],
    ["Name", order.submittedByName || "-"],
    ["Number Items", String(order.orderedCount || 0)],
  ];
  const detail = [
    `<div class="review-detail-meta">${meta.map(([label, value]) => `<div class="review-detail-meta-item"><span>${escape(label)}</span><strong>${escape(value)}</strong></div>`).join("")}</div>`,
  ];
  if (clean(order.notes)) {
    detail.push(`<div class="review-detail-section"><strong>Notes</strong><div class="communication-mobile-message">${escape(order.notes)}</div></div>`);
  }
  if (!selectedItems.length) {
    detail.push('<p class="review-detail-section">No selected items in this order.</p>');
  } else {
    detail.push(`<div class="table-wrap shopping-history-detail-wrap"><table class="shopping-history-table"><thead><tr><th>Category</th><th>Item</th><th>Supplier</th><th>Stored</th><th>Existing Quantity</th><th>Order</th></tr></thead><tbody>${selectedItems.map((item) => `<tr${hexToShoppingRowColor(getShoppingCategoryColor(item.category), 0.15) ? ` style="background:${escape(hexToShoppingRowColor(getShoppingCategoryColor(item.category), 0.15))}"` : ""}><td>${escape(item.category || "-")}</td><td>${escape(item.item || "-")}</td><td>${escape(item.supplier || "-")}</td><td>${escape(item.stored || "-")}</td><td>${escape(item.existingQuantity || "-")}</td><td>Yes</td></tr>`).join("")}</tbody></table></div>`);
  }
  els.shoppingDetailBody.className = "review-detail";
  els.shoppingDetailBody.innerHTML = detail.join("");
  if (els.shoppingCopyOrder) els.shoppingCopyOrder.hidden = !!state.shoppingOpenOrder;
  els.shoppingReopenOrder.hidden = !canReopenShoppingOrderClient(order, state.shoppingHistory);
}

function openShoppingDetailModal(orderId) {
  const order = state.shoppingHistory.find((item) => clean(item.id) === clean(orderId));
  state.shoppingSelectedHistoryId = clean(orderId);
  renderShoppingDetail(order || null);
  els.shoppingDetailModal.hidden = false;
  document.body.classList.add("modal-open");
}

function closeShoppingDetailModal() {
  els.shoppingDetailModal.hidden = true;
  document.body.classList.remove("modal-open");
}

function renderShopping() {
  if (!canApp("shopping")) {
    setShoppingCurrentStatus("Your profile has no access to Shopping.");
    if (els.shoppingOpenRows) els.shoppingOpenRows.innerHTML = '<tr><td colspan="6" class="empty">Your profile has no access to Shopping.</td></tr>';
    if (els.shoppingMobileCards) els.shoppingMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Shopping.</div>';
    if (els.shoppingHistoryRows) els.shoppingHistoryRows.innerHTML = '<tr><td colspan="4" class="empty">Your profile has no access to Shopping.</td></tr>';
    if (els.shoppingHistoryMobileCards) els.shoppingHistoryMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Shopping.</div>';
    return;
  }
  setShoppingTab(state.shoppingTab);
  const order = state.shoppingOpenOrder;
  els.shoppingOpenSummary.textContent = order ? `Open order #${order.orderNumber} · ${order.orderedCount} selected item${order.orderedCount === 1 ? "" : "s"}` : "No open order";
  els.shoppingOpenEmpty.hidden = !!order;
  els.shoppingOpenContent.hidden = !order;
  els.shoppingSaveOrder.hidden = !order;
  els.shoppingNewOrder.hidden = !!order;
  els.shoppingSubmitName.value = state.shoppingSubmitName;
  if (els.shoppingSubmitNotes) els.shoppingSubmitNotes.value = state.shoppingSubmitNotes;
  els.shoppingSubmitNameWrap.hidden = !order || !state.shoppingSubmitPromptOpen;
  if (els.shoppingSubmitOrder) els.shoppingSubmitOrder.textContent = state.shoppingSubmitPromptOpen ? "Confirm Submit" : "Submit Order";
  if (order) {
    renderShoppingCurrentFilters(order);
    renderShoppingCurrentRows(order);
  }
  else {
    if (els.shoppingFilterCategory) els.shoppingFilterCategory.innerHTML = '<option value="">All categories</option>';
    if (els.shoppingFilterStored) els.shoppingFilterStored.innerHTML = '<option value="">All stored</option>';
    if (els.shoppingGroupBy) els.shoppingGroupBy.value = state.shoppingFilters.groupBy;
    if (els.shoppingOpenRows) els.shoppingOpenRows.innerHTML = "";
    if (els.shoppingMobileCards) els.shoppingMobileCards.innerHTML = "";
  }
  renderShoppingHistoryRows();
}

function renderShoppingSettings() {
  if (!els.shoppingSettingsItemsBody) return;
  const settings = state.shoppingSettings || clone(DEFAULT_SHOPPING_SETTINGS);
  if (els.shoppingSettingsEmailRecipients) els.shoppingSettingsEmailRecipients.value = (settings.emailRecipients || []).join("\n");
  if (els.shoppingSettingsWeekdays) {
    els.shoppingSettingsWeekdays.innerHTML = SHOPPING_WEEKDAY_OPTIONS.map((weekday) => `<label class="filter-checkbox"><span>${escape(weekday.label)}</span><input type="checkbox" data-shopping-weekday="${escape(weekday.key)}" ${settings.mandatoryWeekdays.includes(weekday.key) ? "checked" : ""} /></label>`).join("");
  }
  if (els.shoppingSettingsCategoryColors) {
    els.shoppingSettingsCategoryColors.innerHTML = SHOPPING_CATEGORY_OPTIONS.map((category) => `<label class="shopping-category-color-field"><span>${escape(category)}</span><input type="color" data-shopping-category-color="${escape(category)}" value="${escape(settings.categoryColors?.[category] || DEFAULT_SHOPPING_CATEGORY_COLORS[category])}" /></label>`).join("");
  }
  els.shoppingSettingsItemsBody.innerHTML = "";
  settings.items.forEach((item, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><select data-shopping-settings-field="category" data-index="${index}">${SHOPPING_CATEGORY_OPTIONS.map((category) => `<option value="${escape(category)}" ${category === item.category ? "selected" : ""}>${escape(category)}</option>`).join("")}</select></td>
      <td><input data-shopping-settings-field="item" data-index="${index}" value="${escape(item.item)}" /></td>
      <td><input data-shopping-settings-field="supplier" data-index="${index}" value="${escape(item.supplier)}" /></td>
      <td><select data-shopping-settings-field="stored" data-index="${index}"><option value=""></option>${SHOPPING_STORED_OPTIONS.map((stored) => `<option value="${escape(stored)}" ${stored === item.stored ? "selected" : ""}>${escape(stored)}</option>`).join("")}</select></td>
      <td><input data-shopping-settings-field="quantityRequired" data-index="${index}" type="checkbox" ${item.quantityRequired ? "checked" : ""} /></td>
      <td class="row-actions"><button type="button" class="ghost" data-action="remove-shopping-item" data-index="${index}">Remove</button></td>`;
    els.shoppingSettingsItemsBody.appendChild(tr);
  });
}

function addShoppingSettingItem() {
  state.shoppingSettings.items.push({
    id: `shopping-item-${Date.now()}`,
    category: SHOPPING_CATEGORY_OPTIONS[0],
    item: "",
    supplier: "",
    stored: "",
    quantityRequired: true,
  });
  renderShoppingSettings();
}

function onShoppingSettingsInput(event) {
  const index = Number(event.target.dataset.index);
  const field = clean(event.target.dataset.shoppingSettingsField);
  if (field && Number.isFinite(index) && index >= 0) {
    const item = state.shoppingSettings.items[index];
    if (!item) return;
    if (field === "quantityRequired") item.quantityRequired = !!event.target.checked;
    else if (field === "category") item.category = normalizeShoppingCategoryClient(event.target.value);
    else if (field === "stored") item.stored = normalizeShoppingStoredClient(event.target.value);
    else item[field] = clean(event.target.value);
  }
  if (event.target.dataset.shoppingCategoryColor) {
    const category = normalizeShoppingCategoryClient(event.target.dataset.shoppingCategoryColor);
    state.shoppingSettings.categoryColors[category] = normalizeShoppingColorClient(event.target.value, DEFAULT_SHOPPING_CATEGORY_COLORS[category]);
  }
  if (event.target.dataset.shoppingWeekday) {
    const key = clean(event.target.dataset.shoppingWeekday).toLowerCase();
    const set = new Set(state.shoppingSettings.mandatoryWeekdays);
    if (event.target.checked) set.add(key);
    else set.delete(key);
    state.shoppingSettings.mandatoryWeekdays = Array.from(set);
  }
}

function onShoppingSettingsAction(event) {
  const button = event.target.closest("button[data-action]");
  if (button?.dataset.action === "remove-shopping-item") {
    const index = Number(button.dataset.index);
    if (Number.isFinite(index) && index >= 0) {
      state.shoppingSettings.items.splice(index, 1);
      renderShoppingSettings();
    }
    return;
  }
  const weekdayCheckbox = event.target.closest("input[data-shopping-weekday]");
  if (!weekdayCheckbox) return;
  const key = clean(weekdayCheckbox.dataset.shoppingWeekday).toLowerCase();
  const set = new Set(state.shoppingSettings.mandatoryWeekdays);
  if (weekdayCheckbox.checked) set.add(key);
  else set.delete(key);
  state.shoppingSettings.mandatoryWeekdays = Array.from(set);
}

async function saveShoppingSettings() {
  const payload = {
    mandatoryWeekdays: state.shoppingSettings.mandatoryWeekdays,
    emailRecipients: parseEmailList(els.shoppingSettingsEmailRecipients.value),
    categoryColors: { ...state.shoppingSettings.categoryColors },
    items: state.shoppingSettings.items.map((item) => ({
      id: clean(item.id),
      category: normalizeShoppingCategoryClient(item.category),
      item: clean(item.item),
      supplier: clean(item.supplier),
      stored: normalizeShoppingStoredClient(item.stored),
      quantityRequired: !!item.quantityRequired,
    })),
  };
  try {
    const result = await api("/api/shopping-settings", { method: "PUT", body: { settings: payload } });
    state.shoppingSettings = normalizeShoppingSettingsClient(result.settings);
    state.shoppingSettingsLoaded = true;
    renderShoppingSettings();
    renderShopping();
    renderLayout();
    setShoppingSettingsStatus("Shopping configuration saved.");
    showToast("Shopping configuration saved.", "success");
  } catch (e) {
    setShoppingSettingsStatus(`Save failed: ${e.message}`);
    showToast(`Save failed: ${e.message}`, "error");
  }
}

function findShoppingOrderItem(itemId) {
  if (!state.shoppingOpenOrder) return null;
  return state.shoppingOpenOrder.items.find((item) => clean(item.id) === clean(itemId)) || null;
}

function onShoppingOrderInput(event) {
  const itemId = clean(event.target.dataset.shoppingItemId);
  const field = clean(event.target.dataset.shoppingField);
  if (!itemId || !field || !state.shoppingOpenOrder) return;
  const item = findShoppingOrderItem(itemId);
  if (!item) return;
  if (field === "order") {
    item.order = !!event.target.checked;
    state.shoppingOpenOrder.orderedCount = state.shoppingOpenOrder.items.filter((entry) => entry.order).length;
    renderShopping();
    return;
  }
  if (field === "existingQuantity") {
    item.existingQuantity = clean(event.target.value);
  }
}

function onShoppingFilterChange() {
  state.shoppingFilters.category = clean(els.shoppingFilterCategory?.value);
  state.shoppingFilters.stored = clean(els.shoppingFilterStored?.value);
  state.shoppingFilters.groupBy = clean(els.shoppingGroupBy?.value) === "stored" ? "stored" : "category";
  renderShopping();
}

function onShoppingHistoryFilterChange() {
  state.shoppingHistoryFilters.dateFrom = clean(els.shoppingHistoryDateFrom?.value);
  state.shoppingHistoryFilters.dateTo = clean(els.shoppingHistoryDateTo?.value);
  state.shoppingHistoryFilters.name = clean(els.shoppingHistoryName?.value);
  state.shoppingHistoryFilters.category = clean(els.shoppingHistoryCategory?.value);
  state.shoppingHistoryFilters.supplier = clean(els.shoppingHistorySupplier?.value);
  renderShopping();
}

async function createShoppingOrder() {
  try {
    const result = await api("/api/shopping", { method: "POST", body: { action: "create" } });
    state.shoppingOpenOrder = normalizeShoppingOrderClient(result.order, state.shoppingSettings.items);
    state.shoppingLoaded = true;
    state.shoppingSubmitName = "";
    state.shoppingSubmitNotes = "";
    state.shoppingSubmitPromptOpen = false;
    setShoppingTab("current");
    renderShopping();
    renderLayout();
    setShoppingCurrentStatus("Shopping order created.");
    showToast("Shopping order created.", "success");
  } catch (e) {
    setShoppingCurrentStatus(`Create failed: ${e.message}`);
    showToast(`Create failed: ${e.message}`, "error");
  }
}

async function saveShoppingOrderDraft(showSuccess = true) {
  if (!state.shoppingOpenOrder?.id) return;
  try {
    const fallbackOrder = clone(state.shoppingOpenOrder);
    const result = await api(`/api/shopping?id=${encodeURIComponent(state.shoppingOpenOrder.id)}`, {
      method: "PUT",
      body: {
        action: "save",
        items: state.shoppingOpenOrder.items,
      },
    });
    const updatedOrder = normalizeShoppingOrderClient(result.order, state.shoppingSettings.items);
    state.shoppingOpenOrder = updatedOrder?.id ? updatedOrder : fallbackOrder;
    setShoppingTab("current");
    renderShopping();
    if (showSuccess) {
      setShoppingCurrentStatus("Shopping draft saved.");
      showToast("Shopping draft saved.", "success");
    }
  } catch (e) {
    setShoppingCurrentStatus(`Save failed: ${e.message}`);
    showToast(`Save failed: ${e.message}`, "error");
  }
}

async function submitShoppingOrder() {
  if (!state.shoppingOpenOrder?.id) return;
  if (!state.shoppingSubmitPromptOpen) {
    state.shoppingSubmitPromptOpen = true;
    renderShopping();
    setShoppingSubmitStatus("Please enter the name to submit this shopping order.");
    showToast("Please enter the name to submit this shopping order.", "info");
    els.shoppingSubmitName?.focus();
    return;
  }
  try {
    const result = await api(`/api/shopping?id=${encodeURIComponent(state.shoppingOpenOrder.id)}`, {
      method: "PUT",
      body: {
        action: "submit",
        submittedByName: clean(state.shoppingSubmitName || els.shoppingSubmitName.value),
        notes: clean(state.shoppingSubmitNotes || els.shoppingSubmitNotes?.value),
        items: state.shoppingOpenOrder.items,
      },
    });
    const emailError = clean(result.emailResult?.error);
    state.shoppingOpenOrder = null;
    state.shoppingSubmitName = "";
    state.shoppingSubmitNotes = "";
    state.shoppingSubmitPromptOpen = false;
    await loadShoppingData({ silent: true });
    state.shoppingLoaded = true;
    setShoppingTab("history");
    renderShopping();
    renderLayout();
    setShoppingSubmitStatus(emailError ? `Order submitted, but email failed: ${emailError}` : "Shopping order submitted.");
    showToast(emailError ? `Order submitted, but email failed: ${emailError}` : "Shopping order submitted.", emailError ? "error" : "success");
  } catch (e) {
    setShoppingSubmitStatus(`Submit failed: ${e.message}`);
    showToast(`Submit failed: ${e.message}`, "error");
  }
}

function onShoppingHistoryAction(event) {
  const button = event.target.closest("button[data-action]");
  if (button?.dataset.action === "reopen-shopping-order") {
    event.stopPropagation();
    reopenLatestShoppingOrder(button.dataset.id);
    return;
  }
  if (button?.dataset.action === "copy-shopping-order") {
    event.stopPropagation();
    copyShoppingOrderAsDraft(button.dataset.id);
    return;
  }
  const row = event.target.closest("[data-shopping-history-id]");
  if (!row) return;
  openShoppingDetailModal(clean(row.dataset.shoppingHistoryId));
}

async function reopenLatestShoppingOrder(sourceId = "") {
  const latestId = sourceId || state.shoppingHistory[0]?.id;
  if (!latestId) return;
  const order = state.shoppingHistory.find((item) => clean(item.id) === clean(latestId));
  if (!canReopenShoppingOrderClient(order, state.shoppingHistory)) {
    const message = "Only the latest shopping order from today can be reopened.";
    setShoppingDetailStatus(message);
    showToast(message, "error");
    return;
  }
  try {
    const result = await api("/api/shopping", {
      method: "POST",
      body: {
        action: "reopen",
        sourceOrderId: latestId,
      },
    });
    state.shoppingOpenOrder = normalizeShoppingOrderClient(result.order, state.shoppingSettings.items);
    state.shoppingSubmitName = "";
    state.shoppingSubmitNotes = "";
    state.shoppingSubmitPromptOpen = false;
    state.shoppingLoaded = true;
    setShoppingTab("current");
    closeShoppingDetailModal();
    renderShopping();
    renderLayout();
    setShoppingCurrentStatus("Latest shopping order reopened.");
    showToast("Latest shopping order reopened.", "success");
  } catch (e) {
    setShoppingDetailStatus(`Reopen failed: ${e.message}`);
    showToast(`Reopen failed: ${e.message}`, "error");
  }
}

async function copyShoppingOrderAsDraft(sourceId = "") {
  const orderId = sourceId || clean(state.shoppingSelectedHistoryId);
  if (!orderId) return;
  try {
    const result = await api("/api/shopping", {
      method: "POST",
      body: {
        action: "copy",
        sourceOrderId: orderId,
      },
    });
    state.shoppingOpenOrder = normalizeShoppingOrderClient(result.order, state.shoppingSettings.items);
    state.shoppingSubmitName = "";
    state.shoppingSubmitNotes = "";
    state.shoppingSubmitPromptOpen = false;
    state.shoppingLoaded = true;
    setShoppingTab("current");
    closeShoppingDetailModal();
    renderShopping();
    renderLayout();
    setShoppingCurrentStatus("Shopping order copied as a new draft.");
    showToast("Shopping order copied as a new draft.", "success");
  } catch (e) {
    setShoppingDetailStatus(`Copy failed: ${e.message}`);
    showToast(`Copy failed: ${e.message}`, "error");
  }
}

function setBakeryCurrentStatus(text) {
  if (els.bakeryCurrentStatus) els.bakeryCurrentStatus.textContent = text;
}

function setBakerySubmitStatus(text) {
  if (els.bakerySubmitStatus) els.bakerySubmitStatus.textContent = text;
}

function setBakeryHistoryStatus(text) {
  if (els.bakeryHistoryStatus) els.bakeryHistoryStatus.textContent = text;
}

function setBakeryDetailStatus(text) {
  if (els.bakeryDetailStatus) els.bakeryDetailStatus.textContent = text;
}

function setBakerySettingsStatus(text) {
  if (els.bakerySettingsStatus) els.bakerySettingsStatus.textContent = text;
}

async function loadBakerySettings({ silent = false } = {}) {
  try {
    const result = await api("/api/bakery-settings");
    state.bakerySettings = normalizeBakerySettingsClient(result.settings);
    renderBakerySettings();
    if (!silent) setBakerySettingsStatus("Bakery configuration loaded.");
  } catch (e) {
    state.bakerySettings = clone(DEFAULT_BAKERY_SETTINGS);
    renderBakerySettings();
    if (!silent) setBakerySettingsStatus(`Using default bakery settings (${e.message}).`);
  }
}

async function loadBakeryData({ silent = false } = {}) {
  try {
    const result = await api("/api/bakery");
    if (result?.settings) {
      state.bakerySettings = normalizeBakerySettingsClient(result.settings);
      state.bakerySettingsLoaded = true;
    }
    state.bakeryOpenOrder = normalizeBakeryOrderClient(result.openOrder, state.bakerySettings);
    state.bakeryHistory = (Array.isArray(result.history) ? result.history : []).map((order) => normalizeBakeryOrderClient(order, state.bakerySettings));
    state.bakeryLoaded = true;
    renderBakery();
    renderBakerySettings();
    if (!silent) setBakeryCurrentStatus("Bakery orders loaded.");
  } catch (e) {
    state.bakeryOpenOrder = null;
    state.bakeryHistory = [];
    renderBakery();
    if (!silent) {
      setBakeryCurrentStatus(`Failed to load bakery orders: ${e.message}`);
      showToast(`Failed to load bakery orders: ${e.message}`, "error");
    }
  }
}

function setBakeryTab(tab) {
  state.bakeryTab = tab === "history" || tab === "resume" ? tab : "current";
  if (els.bakeryTabCurrent) {
    els.bakeryTabCurrent.classList.toggle("active-tab", state.bakeryTab === "current");
    els.bakeryTabCurrent.classList.toggle("ghost", state.bakeryTab !== "current");
  }
  if (els.bakeryTabHistory) {
    els.bakeryTabHistory.classList.toggle("active-tab", state.bakeryTab === "history");
    els.bakeryTabHistory.classList.toggle("ghost", state.bakeryTab !== "history");
  }
  if (els.bakeryTabResume) {
    els.bakeryTabResume.classList.toggle("active-tab", state.bakeryTab === "resume");
    els.bakeryTabResume.classList.toggle("ghost", state.bakeryTab !== "resume");
  }
  if (els.bakeryPanelCurrent) els.bakeryPanelCurrent.hidden = state.bakeryTab !== "current";
  if (els.bakeryPanelHistory) els.bakeryPanelHistory.hidden = state.bakeryTab !== "history";
  if (els.bakeryPanelResume) els.bakeryPanelResume.hidden = state.bakeryTab !== "resume";
}

function refreshBakeryOpenOrderDerivedState() {
  if (!state.bakeryOpenOrder) return;
  state.bakeryOpenOrder.days = (Array.isArray(state.bakeryOpenOrder.days) ? state.bakeryOpenOrder.days : []).map((day) => normalizeBakeryDayClient(day, state.bakerySettings));
  state.bakeryOpenOrder.generatedText = buildBakeryGeneratedTextClient(state.bakeryOpenOrder, state.bakerySubmitName);
}

function renderBakeryCurrentRows(order) {
  const rows = Array.isArray(order?.days) ? order.days : [];
  if (els.bakeryOpenRows) {
    els.bakeryOpenRows.innerHTML = rows.map((day, index) => `<tr>
      <td>${escape(bakeryDateLabel(day.date))}</td>
      <td><input data-bakery-index="${index}" data-bakery-field="availableBeds" type="number" min="0" step="1" value="${escape(String(day.availableBeds ?? ""))}" /></td>
      <td>${escape(String(day.hostelGuests || 0))}</td>
      <td>${escape(String(day.totalBreads || 0))}</td>
      <td>${escape((day.breadBreakdown || []).map((item) => `${item.name}: ${item.quantity}`).join(" | "))}</td>
      <td><input data-bakery-index="${index}" data-bakery-field="cruzCheckins" type="number" min="0" step="1" value="${escape(String(day.cruzCheckins ?? ""))}" /></td>
      <td>${escape(String(day.pasteisDeNata || 0))}</td>
    </tr>`).join("");
  }
  if (els.bakeryMobileCards) {
    els.bakeryMobileCards.innerHTML = "";
    rows.forEach((day, index) => {
      const card = document.createElement("article");
      card.className = "shopping-mobile-card";
      card.innerHTML = `<div class="shopping-mobile-row">
        <div class="shopping-mobile-main">
          <div class="service-mobile-request">${escape(bakeryDateLabel(day.date))}</div>
          <div class="service-mobile-type">Hostel guests: ${escape(String(day.hostelGuests || 0))} · Breads: ${escape(String(day.totalBreads || 0))}</div>
          <div class="shopping-mobile-supplier">${escape((day.breadBreakdown || []).map((item) => `${item.name}: ${item.quantity}`).join(" | "))}</div>
          <div class="shopping-mobile-supplier">Pastéis de nata: ${escape(String(day.pasteisDeNata || 0))}</div>
        </div>
        <div class="shopping-mobile-inline-order">
          <input data-bakery-index="${index}" data-bakery-field="availableBeds" type="number" min="0" step="1" value="${escape(String(day.availableBeds ?? ""))}" placeholder="Beds" />
          <input class="shopping-mobile-inline-qty" data-bakery-index="${index}" data-bakery-field="cruzCheckins" type="number" min="0" step="1" value="${escape(String(day.cruzCheckins ?? ""))}" placeholder="Cruz" />
        </div>
      </div>`;
      els.bakeryMobileCards.appendChild(card);
    });
  }
}

function renderBakeryCurrentRows(order) {
  const rows = Array.isArray(order?.days) ? order.days : [];
  const breadTypes = bakeryBreadTypeColumnsClient(order);
  const headerRow = document.getElementById("bakery-open-head");
  if (headerRow) {
    headerRow.innerHTML = `<th>Date</th><th class="center-cell">Available Beds (Hostel)</th><th class="center-cell">Check-ins (Cruz)</th>${breadTypes
      .map((breadType) => `<th class="center-cell">${escape(breadType)}</th>`)
      .join("")}<th class="center-cell">Pastéis de Nata</th>`;
  }
  if (els.bakeryOpenRows) {
    els.bakeryOpenRows.innerHTML = rows.map((day, index) => `<tr>
      <td>${escape(bakeryDateLabel(day.date))}<br><small class="bakery-date-meta">Hostel guests: ${escape(day.hostelGuests === "" ? "-" : String(day.hostelGuests))}</small></td>
      <td class="center-cell"><div class="bakery-available-cell"><input data-bakery-index="${index}" data-bakery-field="availableBeds" type="number" min="0" max="${escape(String(state.bakerySettings?.hostelCapacity || 83))}" step="1" value="${escape(day.availableBeds === "" ? "" : String(day.availableBeds))}" required /></div></td>
      <td class="center-cell"><input data-bakery-index="${index}" data-bakery-field="cruzCheckins" type="number" min="0" max="${escape(String(state.bakerySettings?.hostelCapacity || 83))}" step="1" value="${escape(day.cruzCheckins === "" ? "" : String(day.cruzCheckins))}" required /></td>
      ${breadTypes.map((breadType) => `<td class="center-cell">${escape(String(bakeryBreadTypeQuantity(day, breadType) === "" ? "-" : bakeryBreadTypeQuantity(day, breadType)))}</td>`).join("")}
      <td class="center-cell">${escape(String(day.pasteisDeNata === "" ? "-" : day.pasteisDeNata))}</td>
    </tr>`).join("");
  }
  if (els.bakeryMobileCards) {
    els.bakeryMobileCards.innerHTML = "";
    rows.forEach((day, index) => {
      const card = document.createElement("article");
      card.className = "shopping-mobile-card";
      card.innerHTML = `<div class="shopping-mobile-row">
        <div class="shopping-mobile-main">
          <div class="service-mobile-request">${escape(bakeryDateLabel(day.date))}</div>
          <div class="service-mobile-type">Hostel guests: ${escape(day.hostelGuests === "" ? "-" : String(day.hostelGuests))}</div>
          <div class="shopping-mobile-supplier">${escape((day.breadBreakdown || []).map((item) => `${item.name}: ${item.quantity === "" ? "-" : item.quantity}`).join(" | "))}</div>
          <div class="shopping-mobile-supplier">Pastéis de nata: ${escape(String(day.pasteisDeNata === "" ? "-" : day.pasteisDeNata))}</div>
        </div>
        <div class="shopping-mobile-inline-order">
          <input data-bakery-index="${index}" data-bakery-field="availableBeds" type="number" min="0" max="${escape(String(state.bakerySettings?.hostelCapacity || 83))}" step="1" value="${escape(day.availableBeds === "" ? "" : String(day.availableBeds))}" placeholder="Beds" required />
          <input class="shopping-mobile-inline-qty" data-bakery-index="${index}" data-bakery-field="cruzCheckins" type="number" min="0" max="${escape(String(state.bakerySettings?.hostelCapacity || 83))}" step="1" value="${escape(day.cruzCheckins === "" ? "" : String(day.cruzCheckins))}" placeholder="Cruz" required />
        </div>
      </div>`;
      els.bakeryMobileCards.appendChild(card);
    });
  }
}

function bakeryHistoryBreadTypeColumns(orders = []) {
  const seen = new Set();
  const columns = [];
  ((state.bakerySettings?.breadTypes) || []).forEach((item) => {
    const name = clean(item?.name);
    if (!name) return;
    const key = name.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    columns.push(name);
  });
  (Array.isArray(orders) ? orders : []).forEach((order) => {
    (Array.isArray(order?.days) ? order.days : []).forEach((day) => {
      (Array.isArray(day?.breadBreakdown) ? day.breadBreakdown : []).forEach((item) => {
        const name = clean(item?.name);
        if (!name) return;
        const key = name.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        columns.push(name);
      });
    });
  });
  return columns;
}

function bakeryOrderTotals(order) {
  const totals = {
    days: 0,
    pasteisDeNata: 0,
    breadByType: {},
  };
  const days = Array.isArray(order?.days) ? order.days : [];
  totals.days = days.length;
  days.forEach((day) => {
    totals.pasteisDeNata += Number(day?.pasteisDeNata || 0);
    (Array.isArray(day?.breadBreakdown) ? day.breadBreakdown : []).forEach((item) => {
      const name = clean(item?.name);
      if (!name) return;
      totals.breadByType[name] = (totals.breadByType[name] || 0) + Number(item?.quantity || 0);
    });
  });
  return totals;
}

function bakeryResumeMonthKey(order) {
  const raw = clean(order?.submittedAt || order?.updatedAt || order?.createdAt);
  const iso = raw.slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(iso) ? iso.slice(0, 7) : "";
}

function formatBakeryMonthLabel(monthKey) {
  const raw = clean(monthKey);
  if (!/^\d{4}-\d{2}$/.test(raw)) return raw || "-";
  const dt = new Date(`${raw}-01T00:00:00`);
  if (Number.isNaN(dt.getTime())) return raw;
  return new Intl.DateTimeFormat("en-GB", { month: "short", year: "numeric", timeZone: "Europe/Lisbon" }).format(dt);
}

function getBakeryResumeRows() {
  const rows = Array.isArray(state.bakeryHistory) ? state.bakeryHistory : [];
  const breadTypeColumns = bakeryHistoryBreadTypeColumns(rows);
  const buckets = new Map();
  rows.forEach((order) => {
    const monthKey = bakeryResumeMonthKey(order);
    if (!monthKey) return;
    const totals = bakeryOrderTotals(order);
    const current = buckets.get(monthKey) || {
      monthKey,
      pasteisDeNata: 0,
      breadByType: Object.fromEntries(breadTypeColumns.map((name) => [name, 0])),
    };
    current.pasteisDeNata += Number(totals.pasteisDeNata || 0);
    breadTypeColumns.forEach((name) => {
      current.breadByType[name] = (current.breadByType[name] || 0) + Number(totals.breadByType[name] || 0);
    });
    buckets.set(monthKey, current);
  });
  return {
    breadTypeColumns,
    rows: [...buckets.values()].sort((a, b) => clean(b.monthKey).localeCompare(clean(a.monthKey))),
  };
}

function renderBakeryResumeRows() {
  const { breadTypeColumns, rows } = getBakeryResumeRows();
  if (els.bakeryResumeHead) {
    els.bakeryResumeHead.innerHTML = `<th>Month</th><th>Pastéis de nata</th>${breadTypeColumns.map((name) => `<th>${escape(name)}</th>`).join("")}`;
  }
  if (els.bakeryResumeCount) els.bakeryResumeCount.textContent = `${rows.length} month${rows.length === 1 ? "" : "s"}`;
  if (!els.bakeryResumeRows) return;
  if (!rows.length) {
    els.bakeryResumeRows.innerHTML = `<tr><td colspan="${2 + breadTypeColumns.length}" class="empty">No bakery orders found.</td></tr>`;
    return;
  }
  els.bakeryResumeRows.innerHTML = rows.map((row) => `<tr>
      <td>${escape(formatBakeryMonthLabel(row.monthKey))}</td>
      <td>${escape(String(row.pasteisDeNata))}</td>
      ${breadTypeColumns.map((name) => `<td>${escape(String(row.breadByType[name] || 0))}</td>`).join("")}
    </tr>`).join("");
}

function renderBakeryHistoryRows() {
  const rows = Array.isArray(state.bakeryHistory) ? state.bakeryHistory : [];
  const breadTypeColumns = bakeryHistoryBreadTypeColumns(rows);
  const headerRow = els.bakeryHistoryRows?.closest("table")?.querySelector("thead tr");
  if (headerRow) {
    headerRow.innerHTML = `<th>Order Date</th><th>Name</th><th>Days</th><th>Total Pastéis Nata</th>${breadTypeColumns
      .map((name) => `<th>${escape(name)}</th>`)
      .join("")}`;
  }
  if (els.bakeryHistoryCount) els.bakeryHistoryCount.textContent = `${rows.length} order${rows.length === 1 ? "" : "s"}`;
  if (els.bakeryHistoryRows) {
    if (!rows.length) {
      els.bakeryHistoryRows.innerHTML = `<tr><td colspan="${4 + breadTypeColumns.length}" class="empty">No bakery orders found.</td></tr>`;
    } else {
      els.bakeryHistoryRows.innerHTML = rows.map((order) => {
        const totals = bakeryOrderTotals(order);
        return `<tr class="clickable-row" data-bakery-history-id="${escape(order.id)}">
      <td>${escape(formatDateTimeShort(order.submittedAt || order.updatedAt || order.createdAt))}</td>
      <td>${escape(order.submittedByName || "-")}</td>
      <td>${escape(String(totals.days))}</td>
      <td>${escape(String(totals.pasteisDeNata))}</td>
      ${breadTypeColumns.map((name) => `<td>${escape(String(totals.breadByType[name] || 0))}</td>`).join("")}
    </tr>`;
      }).join("");
    }
  }
  if (els.bakeryHistoryMobileCards) {
    els.bakeryHistoryMobileCards.innerHTML = rows.map((order) => {
      const totals = bakeryOrderTotals(order);
      const breadSummary = breadTypeColumns
        .map((name) => `${name}: ${totals.breadByType[name] || 0}`)
        .join(" · ");
      return `<article class="shopping-history-card clickable-row" data-bakery-history-id="${escape(order.id)}">
      <div class="service-mobile-request">${escape(formatDateTimeShort(order.submittedAt || order.updatedAt || order.createdAt))}</div>
      <div class="service-mobile-type">${escape(order.submittedByName || "-")}</div>
      <div class="shopping-mobile-supplier">${escape(String(totals.days))} days · Pastéis: ${escape(String(totals.pasteisDeNata))}</div>
      ${breadSummary ? `<div class="communication-mobile-meta">${escape(breadSummary)}</div>` : ""}
    </article>`;
    }).join("");
  }
}

function renderBakeryDetail(order) {
  if (!els.bakeryDetailBody) return;
  if (els.bakeryDetailResend) els.bakeryDetailResend.hidden = !order;
  if (!order) {
    els.bakeryDetailBody.className = "review-detail empty";
    els.bakeryDetailBody.textContent = "Select an order to see the detail.";
    return;
  }
  els.bakeryDetailBody.className = "review-detail";
  const rows = (Array.isArray(order.days) ? order.days : []).map((day) => `<tr>
    <td>${escape(bakeryDateLabel(day.date))}</td>
    <td>${escape(String(day.availableBeds || 0))}</td>
    <td>${escape(String(day.hostelGuests || 0))}</td>
    <td>${escape(String(day.totalBreads || 0))}</td>
    <td>${escape((day.breadBreakdown || []).map((item) => `${item.name}: ${item.quantity}`).join(" | "))}</td>
    <td>${escape(String(day.cruzCheckins || 0))}</td>
    <td>${escape(String(day.pasteisDeNata || 0))}</td>
  </tr>`).join("");
  els.bakeryDetailBody.innerHTML = `<p><strong>Order #${escape(String(order.orderNumber || ""))}</strong><br>${escape(formatDateTimeShort(order.submittedAt || order.updatedAt || order.createdAt))}<br>${escape(order.submittedByName || "-")}</p>
    <div class="table-wrap shopping-history-detail-wrap"><table class="shopping-history-table"><thead><tr><th>Date</th><th>Available Beds</th><th>Hostel Guests</th><th>Total Breads</th><th>Bread Breakdown</th><th>Cruz</th><th>Pastéis</th></tr></thead><tbody>${rows}</tbody></table></div>
    <label class="settings-email-recipients"><span>Email preview</span><div class="bakery-email-preview">${buildBakeryGeneratedHtmlClient(order, order.submittedByName)}</div></label>`;
}

function openBakeryDetailModal(orderId) {
  state.bakerySelectedHistoryId = clean(orderId);
  const order = (state.bakeryHistory || []).find((item) => item.id === state.bakerySelectedHistoryId) || null;
  setBakeryDetailStatus("");
  renderBakeryDetail(order);
  if (els.bakeryDetailModal) {
    els.bakeryDetailModal.hidden = false;
    document.body.classList.add("modal-open");
  }
}

function closeBakeryDetailModal() {
  if (els.bakeryDetailModal) els.bakeryDetailModal.hidden = true;
  setBakeryDetailStatus("");
  document.body.classList.remove("modal-open");
}

async function resendBakeryOrderEmail() {
  const orderId = clean(state.bakerySelectedHistoryId);
  if (!orderId) return;
  if (els.bakeryDetailResend) els.bakeryDetailResend.disabled = true;
  setBakeryDetailStatus("Resending bakery email...");
  try {
    const result = await api(`/api/bakery?id=${encodeURIComponent(orderId)}`, {
      method: "PUT",
      body: { action: "resend" },
    });
    const emailError = clean(result?.emailResult?.error);
    if (emailError) {
      setBakeryDetailStatus(`Email resend failed: ${emailError}`);
      showToast(`Email resend failed: ${emailError}`, "error");
      return;
    }
    setBakeryDetailStatus("Bakery email resent.");
    showToast("Bakery email resent.", "success");
  } catch (e) {
    setBakeryDetailStatus(`Resend failed: ${e.message}`);
    showToast(`Resend failed: ${e.message}`, "error");
  } finally {
    if (els.bakeryDetailResend) els.bakeryDetailResend.disabled = false;
  }
}

function renderBakery() {
  if (!canApp("bakery")) {
    setBakeryCurrentStatus("Your profile has no access to Bakery.");
    return;
  }
  setBakeryTab(state.bakeryTab);
  const order = state.bakeryOpenOrder;
  const canCreateNewBakeryOrder = !order && isWorkingDayLisbonClient(lisbonTodayIsoClient());
  if (els.bakeryOpenSummary) els.bakeryOpenSummary.textContent = order ? `Open order #${order.orderNumber} · ${bakeryOrderDatesLabel(order)}` : "No open order";
  if (els.bakeryOpenEmpty) els.bakeryOpenEmpty.hidden = !!order;
  if (els.bakeryOpenContent) els.bakeryOpenContent.hidden = !order;
  if (els.bakerySaveOrder) els.bakerySaveOrder.hidden = !order;
  if (els.bakeryNewOrder) els.bakeryNewOrder.hidden = !canCreateNewBakeryOrder;
  if (els.bakerySubmitName) els.bakerySubmitName.value = state.bakerySubmitName;
  if (order) {
    refreshBakeryOpenOrderDerivedState();
    renderBakeryCurrentRows(order);
    if (els.bakeryGeneratedText) els.bakeryGeneratedText.innerHTML = buildBakeryGeneratedHtmlClient(order, state.bakerySubmitName);
  } else {
    if (els.bakeryOpenRows) els.bakeryOpenRows.innerHTML = "";
    if (els.bakeryMobileCards) els.bakeryMobileCards.innerHTML = "";
    if (els.bakeryGeneratedText) els.bakeryGeneratedText.innerHTML = "";
    if (els.bakeryOpenEmpty) {
      els.bakeryOpenEmpty.textContent = canCreateNewBakeryOrder
        ? "No open bakery order. Start a new order to prepare the next bakery email."
        : "New bakery orders can only be created on working days.";
    }
  }
  renderBakeryHistoryRows();
  renderBakeryResumeRows();
}

function setBakerySettingsTab(tab) {
  state.bakerySettingsTab = tab === "types" ? "types" : "table";
  if (els.bakerySettingsTableTab) {
    els.bakerySettingsTableTab.classList.toggle("active-tab", state.bakerySettingsTab === "table");
    els.bakerySettingsTableTab.classList.toggle("ghost", state.bakerySettingsTab !== "table");
  }
  if (els.bakerySettingsTypesTab) {
    els.bakerySettingsTypesTab.classList.toggle("active-tab", state.bakerySettingsTab === "types");
    els.bakerySettingsTypesTab.classList.toggle("ghost", state.bakerySettingsTab !== "types");
  }
  if (els.bakerySettingsTablePanel) els.bakerySettingsTablePanel.hidden = state.bakerySettingsTab !== "table";
  if (els.bakerySettingsTypesPanel) els.bakerySettingsTypesPanel.hidden = state.bakerySettingsTab !== "types";
}

function renderBakerySettings() {
  const settings = state.bakerySettings || clone(DEFAULT_BAKERY_SETTINGS);
  if (els.bakerySelectedBase) els.bakerySelectedBase.value = settings.selectedBase;
  if (els.bakeryHostelCapacity) els.bakeryHostelCapacity.value = settings.hostelCapacity;
  if (els.bakerySettingsEmailRecipients) els.bakerySettingsEmailRecipients.value = (settings.emailRecipients || []).join("\n");
  if (els.bakeryBreadTableBody) {
    els.bakeryBreadTableBody.innerHTML = "";
    settings.breadTable.forEach((row, index) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${escape(String(row.guests))}</td>
        <td><input data-bakery-table-field="baseBaixa" data-index="${index}" type="number" min="0" step="1" value="${escape(String(row.baseBaixa || 0))}" /></td>
        <td><input data-bakery-table-field="baseMedia" data-index="${index}" type="number" min="0" step="1" value="${escape(String(row.baseMedia || 0))}" /></td>
        <td><input data-bakery-table-field="baseAlta" data-index="${index}" type="number" min="0" step="1" value="${escape(String(row.baseAlta || 0))}" /></td>`;
      els.bakeryBreadTableBody.appendChild(tr);
    });
  }
  if (els.bakeryBreadTypesBody) {
    els.bakeryBreadTypesBody.innerHTML = "";
    settings.breadTypes.forEach((item, index) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td><input data-bakery-type-field="name" data-index="${index}" value="${escape(item.name)}" /></td>
        <td><input data-bakery-type-field="percentage" data-index="${index}" type="number" min="0" max="100" step="0.01" value="${escape(String(item.percentage || 0))}" /></td>
        <td class="row-actions"><button type="button" class="ghost" data-action="remove-bakery-type" data-index="${index}">Remove</button></td>`;
      els.bakeryBreadTypesBody.appendChild(tr);
    });
  }
  updateBakeryBreadTypesTotal();
  setBakerySettingsTab(state.bakerySettingsTab);
}

function updateBakeryBreadTypesTotal() {
  if (els.bakeryBreadTypesTotal) {
    const total = ((state.bakerySettings?.breadTypes) || []).reduce((sum, item) => sum + Number(item.percentage || 0), 0);
    els.bakeryBreadTypesTotal.textContent = `Total: ${total}%`;
    els.bakeryBreadTypesTotal.style.color = Math.round(total * 100) === 10000 ? "" : "#b91c1c";
  }
}

function onBakerySettingsInput(event) {
  if (event.target.dataset.bakeryTableField) {
    const idx = Number(event.target.dataset.index);
    const field = clean(event.target.dataset.bakeryTableField);
    const row = state.bakerySettings.breadTable[idx];
    if (row && field) row[field] = Math.max(0, Number.parseInt(event.target.value, 10) || 0);
  }
  if (event.target.dataset.bakeryTypeField) {
    const idx = Number(event.target.dataset.index);
    const field = clean(event.target.dataset.bakeryTypeField);
    const item = state.bakerySettings.breadTypes[idx];
    if (item && field) item[field] = field === "percentage" ? Math.max(0, Number(event.target.value) || 0) : clean(event.target.value);
  }
  if (event.target === els.bakerySelectedBase) state.bakerySettings.selectedBase = normalizeBakeryBaseClient(event.target.value);
  if (event.target === els.bakeryHostelCapacity) state.bakerySettings.hostelCapacity = Math.max(1, Number.parseInt(event.target.value, 10) || 1);
  if (event.target === els.bakerySettingsEmailRecipients) state.bakerySettings.emailRecipients = parseEmailList(event.target.value);
  updateBakeryBreadTypesTotal();
  if (state.bakeryOpenOrder) {
    refreshBakeryOpenOrderDerivedState();
    renderBakery();
  }
}

function onBakerySettingsAction(event) {
  const button = event.target.closest("button[data-action]");
  if (button?.dataset.action === "remove-bakery-type") {
    const index = Number(button.dataset.index);
    if (Number.isFinite(index) && index >= 0) {
      state.bakerySettings.breadTypes.splice(index, 1);
      renderBakerySettings();
    }
  }
}

function addBakeryBreadType() {
  state.bakerySettings.breadTypes.push({ id: `bread-type-${Date.now()}`, name: "", percentage: 0 });
  renderBakerySettings();
}

async function saveBakerySettings() {
  const payload = {
    selectedBase: state.bakerySettings.selectedBase,
    hostelCapacity: state.bakerySettings.hostelCapacity,
    emailRecipients: parseEmailList(els.bakerySettingsEmailRecipients?.value),
    emailConfig: normalizeBakeryEmailConfigClient(state.bakerySettings.emailConfig),
    breadTable: (state.bakerySettings.breadTable || []).map((row) => ({
      guests: row.guests,
      baseBaixa: row.baseBaixa,
      baseMedia: row.baseMedia,
      baseAlta: row.baseAlta,
    })),
    breadTypes: (state.bakerySettings.breadTypes || []).map((item) => ({
      id: clean(item.id),
      name: clean(item.name),
      percentage: Number(item.percentage || 0),
    })).filter((item) => item.name),
  };
  try {
    const result = await api("/api/bakery-settings", { method: "PUT", body: { settings: payload } });
    state.bakerySettings = normalizeBakerySettingsClient(result.settings);
    state.bakerySettingsLoaded = true;
    if (state.bakeryOpenOrder) refreshBakeryOpenOrderDerivedState();
    renderBakerySettings();
    renderBakery();
    setBakerySettingsStatus("Bakery configuration saved.");
    showToast("Bakery configuration saved.", "success");
  } catch (e) {
    setBakerySettingsStatus(`Save failed: ${e.message}`);
    showToast(`Save failed: ${e.message}`, "error");
  }
}

async function createBakeryOrder() {
  if (!isWorkingDayLisbonClient(lisbonTodayIsoClient())) {
    const message = "New bakery orders can only be created on working days.";
    setBakeryCurrentStatus(message);
    showToast(message, "error");
    return;
  }
  try {
    const result = await api("/api/bakery", { method: "POST", body: {} });
    state.bakeryOpenOrder = normalizeBakeryOrderClient(result.order, state.bakerySettings);
    if (state.bakeryOpenOrder) {
      state.bakeryOpenOrder.days = (Array.isArray(state.bakeryOpenOrder.days) ? state.bakeryOpenOrder.days : []).map((day) => normalizeBakeryDayClient({ ...day, availableBeds: "", cruzCheckins: "" }, state.bakerySettings));
      refreshBakeryOpenOrderDerivedState();
    }
    state.bakerySubmitName = "";
    state.bakeryLoaded = true;
    setBakeryTab("current");
    renderBakery();
    setBakeryCurrentStatus("Bakery order created.");
    showToast("Bakery order created.", "success");
  } catch (e) {
    setBakeryCurrentStatus(`Create failed: ${e.message}`);
    showToast(`Create failed: ${e.message}`, "error");
  }
}

async function saveBakeryOrderDraft() {
  if (!state.bakeryOpenOrder?.id) return;
  try {
    const result = await api(`/api/bakery?id=${encodeURIComponent(state.bakeryOpenOrder.id)}`, {
      method: "PUT",
      body: { action: "save", days: state.bakeryOpenOrder.days },
    });
    state.bakeryOpenOrder = normalizeBakeryOrderClient(result.order, state.bakerySettings);
    renderBakery();
    setBakeryCurrentStatus("Bakery draft saved.");
    showToast("Bakery draft saved.", "success");
  } catch (e) {
    setBakeryCurrentStatus(`Save failed: ${e.message}`);
    showToast(`Save failed: ${e.message}`, "error");
  }
}

async function submitBakeryOrder() {
  if (!state.bakeryOpenOrder?.id) return;
  try {
    const result = await api(`/api/bakery?id=${encodeURIComponent(state.bakeryOpenOrder.id)}`, {
      method: "PUT",
      body: { action: "submit", submittedByName: clean(state.bakerySubmitName || els.bakerySubmitName?.value), days: state.bakeryOpenOrder.days },
    });
    const emailError = clean(result.emailResult?.error);
    state.bakeryOpenOrder = null;
    state.bakerySubmitName = "";
    await loadBakeryData({ silent: true });
    state.bakeryLoaded = true;
    setBakeryTab("history");
    renderBakery();
    setBakerySubmitStatus(emailError ? `Order submitted, but email failed: ${emailError}` : "Bakery order submitted.");
    showToast(emailError ? `Order submitted, but email failed: ${emailError}` : "Bakery order submitted.", emailError ? "error" : "success");
  } catch (e) {
    setBakerySubmitStatus(`Submit failed: ${e.message}`);
    showToast(`Submit failed: ${e.message}`, "error");
  }
}

function onBakeryOrderInput(event) {
  const index = Number(event.target.dataset.bakeryIndex);
  const field = clean(event.target.dataset.bakeryField);
  if (!state.bakeryOpenOrder || !Number.isFinite(index) || index < 0 || !field) return;
  const day = state.bakeryOpenOrder.days[index];
  if (!day) return;
  if (field === "availableBeds") day.availableBeds = Math.max(0, Number.parseInt(event.target.value, 10) || 0);
  if (field === "cruzCheckins") day.cruzCheckins = Math.max(0, Number.parseInt(event.target.value, 10) || 0);
  refreshBakeryOpenOrderDerivedState();
  renderBakery();
}

async function saveBakeryOrderDraft() {
  if (!state.bakeryOpenOrder?.id) return;
  const validationError = validateBakeryOrderDays(state.bakeryOpenOrder.days);
  if (validationError) {
    setBakeryCurrentStatus(validationError);
    showToast(validationError, "error");
    return;
  }
  try {
    const result = await api(`/api/bakery?id=${encodeURIComponent(state.bakeryOpenOrder.id)}`, {
      method: "PUT",
      body: { action: "save", days: state.bakeryOpenOrder.days },
    });
    state.bakeryOpenOrder = normalizeBakeryOrderClient(result.order, state.bakerySettings);
    renderBakery();
    setBakeryCurrentStatus("Bakery draft saved.");
    showToast("Bakery draft saved.", "success");
  } catch (e) {
    setBakeryCurrentStatus(`Save failed: ${e.message}`);
    showToast(`Save failed: ${e.message}`, "error");
  }
}

async function submitBakeryOrder() {
  if (!state.bakeryOpenOrder?.id) return;
  const validationError = validateBakeryOrderDays(state.bakeryOpenOrder.days);
  if (validationError) {
    setBakerySubmitStatus(validationError);
    showToast(validationError, "error");
    return;
  }
  try {
    const result = await api(`/api/bakery?id=${encodeURIComponent(state.bakeryOpenOrder.id)}`, {
      method: "PUT",
      body: { action: "submit", submittedByName: clean(state.bakerySubmitName || els.bakerySubmitName?.value), days: state.bakeryOpenOrder.days },
    });
    const emailError = clean(result.emailResult?.error);
    state.bakeryOpenOrder = null;
    state.bakerySubmitName = "";
    await loadBakeryData({ silent: true });
    state.bakeryLoaded = true;
    setBakeryTab("history");
    renderBakery();
    setBakerySubmitStatus(emailError ? `Order submitted, but email failed: ${emailError}` : "Bakery order submitted.");
    showToast(emailError ? `Order submitted, but email failed: ${emailError}` : "Bakery order submitted.", emailError ? "error" : "success");
  } catch (e) {
    setBakerySubmitStatus(`Submit failed: ${e.message}`);
    showToast(`Submit failed: ${e.message}`, "error");
  }
}

function onBakeryOrderInput(event) {
  const index = Number(event.target.dataset.bakeryIndex);
  const field = clean(event.target.dataset.bakeryField);
  if (!state.bakeryOpenOrder || !Number.isFinite(index) || index < 0 || !field) return;
  const day = state.bakeryOpenOrder.days[index];
  if (!day) return;
  const raw = clean(event.target.value);
  const max = Math.max(0, Number(state.bakerySettings?.hostelCapacity || 83));
  if (field === "availableBeds") day.availableBeds = raw ? Math.min(max, Math.max(0, Number.parseInt(raw, 10) || 0)) : "";
  if (field === "cruzCheckins") day.cruzCheckins = raw ? Math.min(max, Math.max(0, Number.parseInt(raw, 10) || 0)) : "";
  refreshBakeryOpenOrderDerivedState();
  renderBakery();
}

function onBakeryHistoryAction(event) {
  const row = event.target.closest("[data-bakery-history-id]");
  if (!row) return;
  openBakeryDetailModal(clean(row.dataset.bakeryHistoryId));
}

function normalizeHoursSettingsClient(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const seen = new Set();
  const people = (Array.isArray(source.people) ? source.people : String(source.people || source.persons || "").split(/[\n,;]/))
    .map((item) => clean(item))
    .filter(Boolean)
    .filter((item) => {
      const key = item.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  return {
    people: people.length ? people : clone(DEFAULT_HOURS_SETTINGS.people),
  };
}

function hoursParseMinutes(value) {
  if (!clean(value)) return null;
  const safe = normalizeTimeInput(value);
  if (!/^\d{2}:\d{2}$/.test(safe)) return null;
  const [hours, minutes] = safe.split(":").map(Number);
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) return null;
  return hours * 60 + minutes;
}

function hoursDurationMinutes(start, finish) {
  const startMinutes = hoursParseMinutes(start);
  const finishMinutes = hoursParseMinutes(finish);
  if (startMinutes == null || finishMinutes == null) return null;
  return finishMinutes - startMinutes;
}

function formatHoursMinutes(minutes) {
  if (minutes == null || minutes <= 0) return "-";
  const totalMinutes = Math.round(Number(minutes) || 0);
  const hours = Math.floor(totalMinutes / 60);
  const remainingMinutes = totalMinutes % 60;
  return `${(totalMinutes / 60).toFixed(2)}h (${hours}h${String(remainingMinutes).padStart(2, "0")}m)`;
}

function formatHoursDuration(start, finish) {
  return formatHoursMinutes(hoursDurationMinutes(start, finish));
}

function emptyHoursDraft() {
  const people = state.hoursSettings?.people || DEFAULT_HOURS_SETTINGS.people;
  return {
    id: "",
    person: people.length === 1 ? people[0] : "",
    date: formatDate(new Date()),
    start: "",
    finish: "",
  };
}

function normalizeHoursTimeValue(value) {
  const raw = clean(value);
  if (!raw) return "";
  if (/^\d{2}:\d{2}$/.test(raw)) return raw;
  if (/^\d{2}:\d{2}:\d{2}$/.test(raw)) return raw.slice(0, 5);
  return raw;
}

function hoursRecordNeedsFinish(record) {
  return !clean(record?.finish);
}

function getPendingHoursRecord(excludeId = "") {
  return state.hoursRecords.find((row) => hoursRecordNeedsFinish(row) && clean(row.id) !== clean(excludeId)) || null;
}

function shouldShowHoursAlert() {
  if (!canApp("hours")) return false;
  return !!getPendingHoursRecord();
}

function normalizeHoursRecordClient(input = {}, settings = state.hoursSettings || DEFAULT_HOURS_SETTINGS) {
  const safeSettings = normalizeHoursSettingsClient(settings);
  const people = safeSettings.people || [];
  const person = clean(input.person);
  return {
    id: clean(input.id),
    person: person || (people.length === 1 ? people[0] : ""),
    date: normalizeDateInput(input.date),
    start: normalizeHoursTimeValue(input.start),
    finish: normalizeHoursTimeValue(input.finish),
    createdAt: clean(input.createdAt || input.created_at),
    updatedAt: clean(input.updatedAt || input.updated_at),
  };
}

function sortHoursRecordsClient(rows) {
  return [...rows].sort((a, b) => {
    const dateCompare = clean(b.date).localeCompare(clean(a.date));
    if (dateCompare !== 0) return dateCompare;
    const personCompare = clean(a.person).localeCompare(clean(b.person));
    if (personCompare !== 0) return personCompare;
    return clean(b.start).localeCompare(clean(a.start));
  });
}

function setHoursStatus(text) {
  if (els.hoursStatus) els.hoursStatus.textContent = text;
}

function setHoursSettingsStatus(text) {
  if (els.hoursSettingsStatus) els.hoursSettingsStatus.textContent = text;
}

function slugifyCashText(value) {
  return clean(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function formatCashDenominationLabel(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount)) return `${value}€`;
  if (Math.abs(amount - Math.round(amount)) < 0.000001) return `${Math.round(amount)}€`;
  return `${String(amount).replace(".", ",")}€`;
}

function normalizeCashMinThresholdsClient(value = {}) {
  const source = value && typeof value === "object" ? value : {};
  return CASH_MIN_ALERT_DENOMINATIONS.reduce((acc, key) => {
    acc[key] = Math.max(0, Math.round(Number(normalizeNumber(source[key]) || 0)));
    return acc;
  }, {});
}

function normalizeCashBoolean(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  return ["true", "1", "yes", "sim", "on"].includes(clean(value).toLowerCase());
}

function normalizeCashSettingsClient(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const shifts = (Array.isArray(source.shifts) ? source.shifts : [])
    .map((shift, index) => ({
      id: clean(shift?.id) || slugifyCashText(`${shift?.name || "shift"}-${index + 1}`),
      name: clean(shift?.name) || `Shift ${index + 1}`,
      startTime: normalizeTimeInput(shift?.startTime ?? shift?.start_time) || "00:00",
    }))
    .filter((shift, index, items) => shift.id && items.findIndex((item) => item.id === shift.id) === index);
  const items = (Array.isArray(source.items) ? source.items : [])
    .map((item, index) => ({
      id: clean(item?.id) || slugifyCashText(`${item?.name || "item"}-${index + 1}`),
      name: clean(item?.name) || `Item ${index + 1}`,
      defaultQuantity: Math.max(0, Math.round(Number(normalizeNumber(item?.defaultQuantity ?? item?.default_quantity) || 0))),
    }))
    .filter((item, index, rows) => item.id && rows.findIndex((row) => row.id === item.id) === index);
  return {
    shifts: shifts.length ? shifts : clone(DEFAULT_CASH_SETTINGS.shifts),
    items: items.length ? items : clone(DEFAULT_CASH_SETTINGS.items),
    minCash: normalizeCashMinThresholdsClient(source.minCash || source.min_cash || DEFAULT_CASH_SETTINGS.minCash),
    maxCashByDenomination: normalizeCashMinThresholdsClient(source.maxCashByDenomination || source.max_cash_by_denomination || source.maxCash || source.max_cash || DEFAULT_CASH_SETTINGS.maxCashByDenomination),
    minimumCashEmailEnabled: normalizeCashBoolean(source.minimumCashEmailEnabled ?? source.minimum_cash_email_enabled),
    maximumCashEmailEnabled: normalizeCashBoolean(source.maximumCashEmailEnabled ?? source.maximum_cash_email_enabled),
    maximumCash: Number((normalizeNumber(source.maximumCash ?? source.maximum_cash) || 0).toFixed(2)),
    managerAlertEmails: parseEmailList(source.managerAlertEmails || source.manager_alert_emails || source.managerAlertEmail || source.manager_alert_email),
  };
}

function cashShiftById(id, settings = state.cashSettings) {
  return (settings?.shifts || []).find((shift) => clean(shift.id) === clean(id)) || null;
}

function cashShiftOrder(settings = state.cashSettings) {
  return (settings?.shifts || DEFAULT_CASH_SETTINGS.shifts).map((shift) => clean(shift.id));
}

function cashShiftSortIndex(value, settings = state.cashSettings) {
  const raw = clean(value);
  const order = cashShiftOrder(settings);
  if (raw) {
    const directIndex = order.indexOf(raw);
    if (directIndex >= 0) return directIndex;
  }
  const normalized = cashShiftDisplayLabel(raw).toLowerCase();
  if (normalized === "n") return 0;
  if (normalized === "m") return 1;
  if (normalized === "t") return 2;
  return Number.MAX_SAFE_INTEGER;
}

function normalizeCashCountsClient(value = {}) {
  const source = value && typeof value === "object" ? value : {};
  return CASH_DENOMINATIONS.reduce((acc, denom) => {
    const numeric = normalizeNumber(source[denom.key]);
    acc[denom.key] = numeric == null ? 0 : Math.max(0, Math.round(Number(numeric || 0)));
    return acc;
  }, {});
}

function normalizeCashMoneyText(value) {
  const raw = clean(value).replace(",", ".");
  if (!raw) return "";
  if (/^-?\d*(?:\.\d{0,2})?$/.test(raw)) return raw;
  const numeric = normalizeNumber(raw);
  return numeric == null ? "" : String(Number(numeric.toFixed(2)));
}

function cashMoneyNumber(value) {
  const numeric = normalizeNumber(value);
  if (numeric == null) return 0;
  return Number(numeric.toFixed(2));
}

function normalizeCashStatusClient(value, fallback = "C") {
  const raw = clean(value).toUpperCase();
  if (raw === "O") return "O";
  if (raw === "C") return "C";
  return fallback === "O" ? "O" : "C";
}

function isCashOpenStatusClient(value) {
  return normalizeCashStatusClient(value) === "O";
}

function normalizeCashItemCountsClient(value = {}, settings = state.cashSettings) {
  const source = value && typeof value === "object" ? value : {};
  return (settings?.items || []).reduce((acc, item) => {
    const sourceKey = Object.keys(source).find((key) => slugifyCashText(key) === slugifyCashText(item.name));
    const raw = source[item.id] ?? source[item.name] ?? (sourceKey ? source[sourceKey] : undefined);
    const numeric = normalizeNumber(raw);
    acc[item.id] = numeric == null ? null : Math.max(0, Math.round(Number(numeric || 0)));
    return acc;
  }, {});
}

function normalizeCashItemJustificationsClient(value = {}, settings = state.cashSettings) {
  const source = value && typeof value === "object" ? value : {};
  return (settings?.items || []).reduce((acc, item) => {
    const sourceKey = Object.keys(source).find((key) => slugifyCashText(key) === slugifyCashText(item.name));
    acc[item.id] = clean(source[item.id] ?? source[item.name] ?? (sourceKey ? source[sourceKey] : ""));
    return acc;
  }, {});
}

function normalizeCashRecordClient(input = {}, settings = state.cashSettings) {
  const safeSettings = normalizeCashSettingsClient(settings);
  const shiftNameRaw = clean(input.shiftName ?? input.shift ?? input.shift_name);
  const shiftId = clean(input.shiftId ?? input.shift_id)
    || clean(safeSettings.shifts.find((item) => clean(item.name).toLowerCase() === shiftNameRaw.toLowerCase())?.id);
  const shift = cashShiftById(shiftId, safeSettings);
  return {
    id: clean(input.id),
    day: clean(input.day ?? input.date),
    shiftId: clean(shift?.id || shiftId),
    shiftName: clean(shift?.name || shiftNameRaw),
    status: normalizeCashStatusClient(input.status, "C"),
    name: clean(input.name),
    denominations: normalizeCashCountsClient(input.denominations),
    cardPos: normalizeCashMoneyText(input.cardPos ?? input.card_pos),
    cashFdm: normalizeCashMoneyText(input.cashFdm ?? input.cash_fdm),
    cardFdm: normalizeCashMoneyText(input.cardFdm ?? input.card_fdm),
    justification: clean(input.justification),
    itemCounts: normalizeCashItemCountsClient(input.itemCounts ?? input.item_counts, safeSettings),
    itemJustifications: normalizeCashItemJustificationsClient(input.itemJustifications ?? input.item_justifications, safeSettings),
    createdAt: clean(input.createdAt ?? input.created_at),
    updatedAt: clean(input.updatedAt ?? input.updated_at),
  };
}

function getOpenCashRecordClient(records = state.cashRecords) {
  return (records || []).find((record) => isCashOpenStatusClient(record?.status)) || null;
}

function getLastClosedCashRecordClient(records = state.cashRecords, settings = state.cashSettings) {
  const closed = cashSortRecordsClient(
    (records || []).filter((record) => !isCashOpenStatusClient(record?.status)),
    settings
  );
  return closed.at(-1) || null;
}

function cashSortRecordsClient(records = state.cashRecords, settings = state.cashSettings) {
  return [...records].sort((a, b) => {
    const dayCompare = clean(a.day).localeCompare(clean(b.day));
    if (dayCompare !== 0) return dayCompare;
    const shiftCompare = cashShiftSortIndex(clean(a.shiftId) || clean(a.shiftName), settings) - cashShiftSortIndex(clean(b.shiftId) || clean(b.shiftName), settings);
    if (shiftCompare !== 0) return shiftCompare;
    return clean(a.shiftName).localeCompare(clean(b.shiftName));
  });
}

function cashRecordKey(record) {
  return `${clean(record?.day)}::${clean(record?.shiftId)}`;
}

function calculateCashTotalClient(denominations = {}) {
  const total = CASH_DENOMINATIONS.reduce((sum, denom) => sum + Number(denominations?.[denom.key] || 0) * denom.value, 0);
  return Number(total.toFixed(2));
}

function shiftCashDay(value, days) {
  const raw = clean(value);
  if (!raw) return "";
  const date = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(date.getTime())) return raw;
  date.setDate(date.getDate() + days);
  return formatDate(date);
}

function getPreviousCashDescriptor(day, shiftId, settings = state.cashSettings) {
  const shifts = settings?.shifts || DEFAULT_CASH_SETTINGS.shifts;
  const index = shifts.findIndex((item) => clean(item.id) === clean(shiftId));
  if (index === -1) return null;
  if (index > 0) return { day, shiftId: shifts[index - 1].id };
  return { day: shiftCashDay(day, -1), shiftId: shifts[shifts.length - 1].id };
}

function getNextCashDescriptor(day, shiftId, settings = state.cashSettings) {
  const shifts = settings?.shifts || DEFAULT_CASH_SETTINGS.shifts;
  const index = shifts.findIndex((item) => clean(item.id) === clean(shiftId));
  if (index === -1) return null;
  if (index < shifts.length - 1) return { day, shiftId: shifts[index + 1].id };
  return { day: shiftCashDay(day, 1), shiftId: shifts[0].id };
}

function getPreviousCashRecordClient(record, records = state.cashRecords, settings = state.cashSettings) {
  const previousRef = getPreviousCashDescriptor(record?.day, record?.shiftId, settings);
  if (!previousRef) return null;
  return (records || []).find((row) => clean(row.id) !== clean(record?.id) && clean(row.day) === clean(previousRef.day) && clean(row.shiftId) === clean(previousRef.shiftId)) || null;
}

function buildComputedCashRowsClient(records = state.cashRecords, settings = state.cashSettings) {
  const sorted = cashSortRecordsClient(records, settings);
  const byKey = new Map(sorted.map((row) => [cashRecordKey(row), row]));
  return sorted.map((row) => {
    const previousRef = getPreviousCashDescriptor(row.day, row.shiftId, settings);
    const previous = previousRef ? byKey.get(`${previousRef.day}::${previousRef.shiftId}`) : null;
    const cashTotal = calculateCashTotalClient(row.denominations);
    const calculatedCash = previous ? Number((calculateCashTotalClient(previous.denominations) + cashMoneyNumber(row.cashFdm)).toFixed(2)) : null;
    const diffCash = calculatedCash == null ? null : Number((cashTotal - calculatedCash).toFixed(2));
    const diffCard = Number((cashMoneyNumber(row.cardPos) - cashMoneyNumber(row.cardFdm)).toFixed(2));
    const itemDiffs = (settings?.items || []).map((item) => {
      const counted = row.itemCounts?.[item.id];
      const diff = counted == null ? null : counted - Number(item.defaultQuantity || 0);
      return { itemId: item.id, counted, diff, defaultQuantity: Number(item.defaultQuantity || 0) };
    });
    return {
      ...row,
      cashTotal,
      calculatedCash,
      diffCash,
      diffCard,
      itemDiffs,
      hasItemDiffs: itemDiffs.some((item) => item.diff != null && item.diff !== 0),
      hasIncompleteItemData: itemDiffs.some((item) => item.counted == null),
    };
  });
}

function getNextExpectedCashRecordClient(records = state.cashRecords, settings = state.cashSettings) {
  const shifts = settings?.shifts || DEFAULT_CASH_SETTINGS.shifts;
  const sorted = cashSortRecordsClient(records, settings);
  const last = sorted.at(-1);
  if (!last) {
    return { day: lisbonTodayIsoClient(), shiftId: shifts[0].id, shiftName: shifts[0].name };
  }
  const nextRef = getNextCashDescriptor(last.day, last.shiftId, settings);
  const shift = cashShiftById(nextRef?.shiftId, settings) || shifts[0];
  return { day: nextRef?.day || last.day, shiftId: shift.id, shiftName: shift.name };
}

function emptyCashDraft() {
  const next = getNextExpectedCashRecordClient(state.cashRecords, state.cashSettings);
  const source = getLastClosedCashRecordClient(state.cashRecords, state.cashSettings);
  return normalizeCashRecordClient({
    id: "",
    day: next.day,
    shiftId: next.shiftId,
    shiftName: next.shiftName,
    status: "O",
    name: "",
    denominations: clone(source?.denominations || {}),
    cardPos: "",
    cashFdm: "",
    cardFdm: "",
    justification: "",
    itemCounts: (state.cashSettings?.items || DEFAULT_CASH_SETTINGS.items).reduce((acc, item) => {
      const sourceCounts = source?.itemCounts || {};
      const sourceValue = sourceCounts[item.id];
      acc[item.id] = sourceValue == null ? null : sourceValue;
      return acc;
    }, {}),
    itemJustifications: {},
  }, state.cashSettings);
}

function hasCashDraft() {
  const draft = state.cashEditingId ? state.cashEditDraft : (state.cashOpenDraft || state.cashDraft);
  return !!(
    clean(draft?.name) ||
    clean(draft?.day) ||
    Object.values(draft?.denominations || {}).some((value) => Number(value || 0) !== 0) ||
    clean(draft?.cardPos) ||
    clean(draft?.cashFdm) ||
    clean(draft?.cardFdm)
  );
}

function cashDraftComputed(draft) {
  const rows = [...state.cashRecords.filter((row) => clean(row.id) !== clean(draft.id)), draft];
  return buildComputedCashRowsClient(rows, state.cashSettings).find((row) => clean(row.id) === clean(draft.id)) || {
    ...draft,
    cashTotal: calculateCashTotalClient(draft.denominations),
    calculatedCash: null,
    diffCash: null,
    diffCard: Number((cashMoneyNumber(draft.cardPos) - cashMoneyNumber(draft.cardFdm)).toFixed(2)),
    itemDiffs: [],
    hasItemDiffs: false,
    hasIncompleteItemData: (state.cashSettings?.items || []).some((item) => draft.itemCounts?.[item.id] == null),
  };
}

function validateCashDraftClient(draft, { isCreate = false } = {}) {
  if (!clean(draft?.name)) return "Name is required.";
  const duplicate = state.cashRecords.find((row) => cashRecordKey(row) === cashRecordKey(draft) && clean(row.id) !== clean(draft.id));
  if (duplicate) return `A cash control record for ${draft.day} ${draft.shiftName || draft.shiftId} already exists.`;
  const existing = state.cashRecords.find((row) => clean(row.id) === clean(draft.id)) || null;
  const otherOpen = state.cashRecords.find((row) => isCashOpenStatusClient(row.status) && clean(row.id) !== clean(draft.id));
  if (isCreate && otherOpen) return "Close the current open shift before adding a new record.";
  if (isCashOpenStatusClient(draft.status) && otherOpen) return "Only one cash control shift can stay open at a time.";
  if (existing && normalizeCashStatusClient(existing.status) === "C" && normalizeCashStatusClient(draft.status) !== "C") {
    return "A closed cash control shift cannot be reopened.";
  }
  if (isCreate) {
    const next = getNextExpectedCashRecordClient();
    if (clean(draft.day) !== clean(next.day) || clean(draft.shiftId) !== clean(next.shiftId)) {
      return `The next record must be ${next.day} ${next.shiftName}.`;
    }
  }
  if (isCashOpenStatusClient(draft.status)) return "";
  for (const item of state.cashSettings?.items || []) {
    const counted = draft.itemCounts?.[item.id];
    if (counted == null || counted === "") return `Count is required for item ${item.name}.`;
    if (Number(counted) !== Number(item.defaultQuantity || 0) && !clean(draft.itemJustifications?.[item.id])) {
      return `Justification is required for item ${item.name}.`;
    }
  }
  const computed = cashDraftComputed(draft);
  if ((computed.diffCash != null && computed.diffCash !== 0) || computed.diffCard !== 0) {
    if (!clean(draft.justification)) return "Justification is required when Dif. Cash or Dif. Card is not zero.";
  }
  return "";
}

function formatCashMoney(value) {
  if (value == null || value === "") return "-";
  return formatMoney(cashMoneyNumber(value));
}

function cashItemDiffLabel(record) {
  return "Items";
}

function cashItemsAlertClass(record) {
  return record.hasItemDiffs || record.hasIncompleteItemData ? " cash-items-alert" : "";
}

function cashDefaultDateFrom() {
  const today = lisbonTodayIsoClient();
  const [year, month] = today.split("-").map(Number);
  const date = new Date(Date.UTC(year, (month || 1) - 1, 1));
  date.setUTCMonth(date.getUTCMonth() - 3);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-01`;
}

function cashShiftDisplayLabel(value) {
  const raw = clean(value).toLowerCase();
  if (raw === "night" || raw === "n") return "N";
  if (raw === "morning" || raw === "m") return "M";
  if (raw === "afternoon" || raw === "t") return "T";
  return clean(value);
}

function formatCashDateCompact(value) {
  const raw = clean(value);
  if (!raw) return "-";
  const dt = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(dt.getTime())) return raw;
  const day = String(dt.getDate()).padStart(2, "0");
  const month = new Intl.DateTimeFormat("en-GB", { month: "short", timeZone: "Europe/Lisbon" }).format(dt);
  return `${day}${month}`;
}

function filteredCashRows(rows) {
  const dateFrom = clean(state.cashFilters.dateFrom);
  const dateTo = clean(state.cashFilters.dateTo);
  const shift = clean(state.cashFilters.shift).toLowerCase();
  const name = clean(state.cashFilters.name).toLowerCase();
  return rows.filter((row) => {
    const day = clean(row.day);
    if (dateFrom && day && day < dateFrom) return false;
    if (dateTo && day && day > dateTo) return false;
    if (shift) {
      const rowShift = clean(row.shiftId || row.shiftName).toLowerCase();
      const rowShiftLabel = cashShiftDisplayLabel(row.shiftName || row.shiftId || "").toLowerCase();
      if (rowShift !== shift && rowShiftLabel !== shift) return false;
    }
    if (name && !clean(row.name).toLowerCase().includes(name)) return false;
    return true;
  });
}

function visibleCashRows(rows, settings = state.cashSettings) {
  const filtered = filteredCashRows(rows);
  return [...filtered].sort((a, b) => {
    const dayCompare = clean(b.day).localeCompare(clean(a.day));
    if (dayCompare !== 0) return dayCompare;
    const shiftCompare = cashShiftSortIndex(clean(b.shiftId) || clean(b.shiftName), settings) - cashShiftSortIndex(clean(a.shiftId) || clean(a.shiftName), settings);
    if (shiftCompare !== 0) return shiftCompare;
    return clean(b.shiftName).localeCompare(clean(a.shiftName));
  });
}

function renderCashScreenTabs() {
  const canSeeResume = isAdministratorProfile();
  if (!canSeeResume && state.cashScreen === "resume") {
    state.cashScreen = "list";
  }
  if (els.cashTabList) {
    els.cashTabList.classList.toggle("active-tab", state.cashScreen === "list");
    els.cashTabList.classList.toggle("ghost", state.cashScreen !== "list");
  }
  if (els.cashTabDetail) {
    els.cashTabDetail.classList.toggle("active-tab", state.cashScreen === "detail");
    els.cashTabDetail.classList.toggle("ghost", state.cashScreen !== "detail");
  }
  if (els.cashTabItems) {
    els.cashTabItems.classList.toggle("active-tab", state.cashScreen === "items");
    els.cashTabItems.classList.toggle("ghost", state.cashScreen !== "items");
  }
  if (els.cashTabResume) {
    els.cashTabResume.hidden = !canSeeResume;
    els.cashTabResume.classList.toggle("active-tab", state.cashScreen === "resume");
    els.cashTabResume.classList.toggle("ghost", state.cashScreen !== "resume");
  }
  if (els.cashPanelList) els.cashPanelList.hidden = state.cashScreen !== "list";
  if (els.cashPanelDetail) els.cashPanelDetail.hidden = state.cashScreen !== "detail";
  if (els.cashPanelItems) els.cashPanelItems.hidden = state.cashScreen !== "items";
  if (els.cashPanelResume) els.cashPanelResume.hidden = state.cashScreen !== "resume" || !canSeeResume;
}

function renderCashDetailRows(rows) {
  if (!els.cashDetailRows) return;
  els.cashDetailRows.innerHTML = "";
  if (!rows.length) {
    els.cashDetailRows.innerHTML = '<tr><td colspan="19" class="empty">No cash records found.</td></tr>';
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(formatCashDateCompact(row.day))}</td>
      <td>${escape(cashShiftDisplayLabel(row.shiftName || row.shiftId || ""))}</td>
      <td>${escape(row.name || "-")}</td>
      ${CASH_DENOMINATIONS.map((denom) => `<td>${escape(String(Number(row.denominations?.[denom.key] || 0)))}</td>`).join("")}
      <td>${escape(formatCashMoney(row.cashTotal))}</td>`;
    els.cashDetailRows.appendChild(tr);
  });
}

function renderCashItemDetailRows(rows) {
  if (!els.cashItemDetailHead || !els.cashItemDetailRows) return;
  const items = state.cashSettings?.items || DEFAULT_CASH_SETTINGS.items;
  els.cashItemDetailHead.innerHTML = `<tr>
    <th>Day</th>
    <th>Shift</th>
    <th>Name</th>
    ${items.map((item) => `<th>${escape(item.name)}<br />(${escape(String(item.defaultQuantity ?? 0))})</th>`).join("")}
  </tr>`;
  els.cashItemDetailRows.innerHTML = "";
  if (!rows.length) {
    els.cashItemDetailRows.innerHTML = `<tr><td colspan="${3 + items.length}" class="empty">No cash records found.</td></tr>`;
  } else {
    rows.forEach((row) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${escape(formatCashDateCompact(row.day))}</td>
        <td>${escape(cashShiftDisplayLabel(row.shiftName || row.shiftId || ""))}</td>
        <td>${escape(row.name || "-")}</td>
        ${items.map((item) => {
          const value = row.itemCounts?.[item.id];
          return `<td>${value == null ? "" : escape(String(value))}</td>`;
      }).join("")}`;
      els.cashItemDetailRows.appendChild(tr);
    });
  }
}

function renderCashResumeRows(rows) {
  if (!els.cashResumeRows) return;
  els.cashResumeRows.innerHTML = "";
  const buckets = new Map();
  rows.forEach((row) => {
    const month = clean(row.day).slice(0, 7);
    if (!month) return;
    const current = buckets.get(month) || {
      month,
      cashFdm: 0,
      cardFdm: 0,
      diffCash: 0,
      diffCard: 0,
      absDiffCash: 0,
      absDiffCard: 0,
    };
    current.cashFdm = Number((current.cashFdm + cashMoneyNumber(row.cashFdm)).toFixed(2));
    current.cardFdm = Number((current.cardFdm + cashMoneyNumber(row.cardFdm)).toFixed(2));
    current.diffCash = Number((current.diffCash + cashMoneyNumber(row.diffCash)).toFixed(2));
    current.diffCard = Number((current.diffCard + cashMoneyNumber(row.diffCard)).toFixed(2));
    current.absDiffCash = Number((current.absDiffCash + Math.abs(cashMoneyNumber(row.diffCash))).toFixed(2));
    current.absDiffCard = Number((current.absDiffCard + Math.abs(cashMoneyNumber(row.diffCard))).toFixed(2));
    buckets.set(month, current);
  });
  const summaryRows = [...buckets.values()].sort((a, b) => clean(b.month).localeCompare(clean(a.month)));
  if (!summaryRows.length) {
    els.cashResumeRows.innerHTML = '<tr><td colspan="7" class="empty">No cash records found.</td></tr>';
    return;
  }
  summaryRows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(row.month)}</td>
      <td>${escape(formatCashMoney(row.cashFdm))}</td>
      <td>${escape(formatCashMoney(row.cardFdm))}</td>
      <td class="${cashDiffValueClass(row.diffCash)}">${escape(formatCashDiffValue(row.diffCash))}</td>
      <td class="${cashDiffValueClass(row.diffCard)}">${escape(formatCashDiffValue(row.diffCard))}</td>
      <td>${escape(formatCashMoney(row.absDiffCash))}</td>
      <td>${escape(formatCashMoney(row.absDiffCard))}</td>`;
    els.cashResumeRows.appendChild(tr);
  });
}

function cashSummaryButtonLabel(record) {
  return record.cashTotal > 0 ? formatCashMoney(record.cashTotal) : "Cash";
}

function cashDiffValueClass(value) {
  return value != null && Number(value) !== 0 ? "cash-diff-value" : "";
}

function formatCashDiffValue(value) {
  return value != null && Number(value) === 0 ? "" : formatCashMoney(value);
}

function buildCashInlineRow() {
  const draft = state.cashDraft || emptyCashDraft();
  const computed = cashDraftComputed(draft);
  const tr = document.createElement("tr");
  tr.className = "cash-inline-row";
  tr.innerHTML = `<td>${escape(formatCashDateCompact(draft.day))}</td>
    <td>${escape(cashShiftDisplayLabel(draft.shiftName || cashShiftById(draft.shiftId)?.name || ""))}</td>
    <td><input data-cash-field="name" data-cash-surface="table" data-scope="new" type="text" value="${escape(draft.name)}" /></td>
    <td><button type="button" class="ghost" data-cash-action="cash" data-scope="new">${escape(cashSummaryButtonLabel(computed))}</button></td>
    <td><input class="cash-money-input" data-cash-field="cardPos" data-cash-surface="table" data-scope="new" type="text" inputmode="decimal" value="${escape(String(draft.cardPos ?? ""))}" /></td>
    <td><input class="cash-money-input" data-cash-field="cashFdm" data-cash-surface="table" data-scope="new" type="text" inputmode="decimal" value="${escape(String(draft.cashFdm ?? ""))}" /></td>
    <td><input class="cash-money-input" data-cash-field="cardFdm" data-cash-surface="table" data-scope="new" type="text" inputmode="decimal" value="${escape(String(draft.cardFdm ?? ""))}" /></td>
    <td>${escape(formatCashMoney(computed.calculatedCash))}</td>
    <td>${escape(formatCashMoney(computed.diffCash))}</td>
    <td>${escape(formatCashMoney(computed.diffCard))}</td>
    <td><input data-cash-field="justification" data-cash-surface="table" data-scope="new" type="text" value="${escape(draft.justification)}" /></td>
    <td><button type="button" class="ghost${cashItemsAlertClass(computed)}" data-cash-action="items" data-scope="new">${escape(cashItemDiffLabel(computed))}</button></td>
    <td><button type="button" data-cash-action="save-new">Add</button></td>`;
  return tr;
}

function buildCashReadOnlyRow(record) {
  const tr = document.createElement("tr");
  if ((record.diffCash != null && record.diffCash !== 0) || record.diffCard !== 0) tr.classList.add("cash-diff-row");
  tr.innerHTML = `<td>${escape(formatCashDateCompact(record.day))}</td>
    <td>${escape(cashShiftDisplayLabel(record.shiftName || ""))}</td>
    <td>${escape(record.name || "-")}</td>
    <td><button type="button" class="ghost" data-cash-action="cash-existing" data-id="${escape(record.id)}">${escape(cashSummaryButtonLabel(record))}</button></td>
    <td>${escape(formatCashMoney(record.cardPos))}</td>
    <td>${escape(formatCashMoney(record.cashFdm))}</td>
    <td>${escape(formatCashMoney(record.cardFdm))}</td>
    <td>${escape(formatCashMoney(record.calculatedCash))}</td>
    <td class="${cashDiffValueClass(record.diffCash)}">${escape(formatCashDiffValue(record.diffCash))}</td>
    <td class="${cashDiffValueClass(record.diffCard)}">${escape(formatCashDiffValue(record.diffCard))}</td>
    <td>${escape(record.justification || "-")}</td>
    <td><button type="button" class="ghost${cashItemsAlertClass(record)}" data-cash-action="items-existing" data-id="${escape(record.id)}">${escape(cashItemDiffLabel(record))}</button></td>
    <td><button type="button" class="ghost" data-cash-action="edit" data-id="${escape(record.id)}">Edit</button></td>`;
  return tr;
}

function buildCashEditableRow(record, { openMode = false } = {}) {
  const scope = openMode ? "open" : "edit";
  const draft = currentCashDraft(scope, record.id) || record;
  const computed = cashDraftComputed(draft);
  const tr = document.createElement("tr");
  tr.className = "cash-inline-row";
  tr.innerHTML = `<td>${escape(formatCashDateCompact(draft.day))}</td>
    <td>${escape(cashShiftDisplayLabel(draft.shiftName || ""))}</td>
    <td><input data-cash-field="name" data-cash-surface="table" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" value="${escape(draft.name)}" /></td>
    <td><button type="button" class="ghost" data-cash-action="cash" data-id="${escape(record.id)}" data-scope="${escape(scope)}">${escape(cashSummaryButtonLabel(computed))}</button></td>
    <td><input class="cash-money-input" data-cash-field="cardPos" data-cash-surface="table" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" inputmode="decimal" value="${escape(String(draft.cardPos ?? ""))}" /></td>
    <td><input class="cash-money-input" data-cash-field="cashFdm" data-cash-surface="table" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" inputmode="decimal" value="${escape(String(draft.cashFdm ?? ""))}" /></td>
    <td><input class="cash-money-input" data-cash-field="cardFdm" data-cash-surface="table" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" inputmode="decimal" value="${escape(String(draft.cardFdm ?? ""))}" /></td>
    <td>${escape(formatCashMoney(computed.calculatedCash))}</td>
    <td class="${cashDiffValueClass(computed.diffCash)}">${escape(formatCashDiffValue(computed.diffCash))}</td>
    <td class="${cashDiffValueClass(computed.diffCard)}">${escape(formatCashDiffValue(computed.diffCard))}</td>
    <td><input data-cash-field="justification" data-cash-surface="table" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" value="${escape(draft.justification)}" /></td>
    <td><button type="button" class="ghost${cashItemsAlertClass(computed)}" data-cash-action="items" data-id="${escape(record.id)}" data-scope="${escape(scope)}">${escape(cashItemDiffLabel(computed))}</button></td>
    <td>${openMode
      ? `<button type="button" data-cash-action="save-edit" data-scope="${escape(scope)}" data-id="${escape(record.id)}">Save</button> <button type="button" class="ghost" data-cash-action="close-open" data-id="${escape(record.id)}">Close</button>`
      : `<button type="button" data-cash-action="save-edit" data-scope="${escape(scope)}" data-id="${escape(record.id)}">Save</button> <button type="button" class="ghost" data-cash-action="cancel-edit" data-id="${escape(record.id)}">Cancel</button>`}</td>`;
  return tr;
}

function buildCashReadOnlyCard(record) {
  const card = document.createElement("article");
  card.className = "cash-mobile-card";
  if ((record.diffCash != null && record.diffCash !== 0) || record.diffCard !== 0) card.classList.add("cash-diff-card");
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(formatCashDateCompact(record.day))} · ${escape(cashShiftDisplayLabel(record.shiftName || ""))}</div>
        <div class="communication-mobile-meta">${escape(record.name || "-")}</div>
      </div>
    </div>
    <div class="communication-mobile-grid">
      <div class="communication-mobile-field"><small>Cash</small><div class="communication-mobile-message"><button type="button" class="ghost" data-cash-action="cash-existing" data-id="${escape(record.id)}">${escape(cashSummaryButtonLabel(record))}</button></div></div>
      <div class="communication-mobile-field"><small>Card POS</small><div class="communication-mobile-message">${escape(formatCashMoney(record.cardPos))}</div></div>
      <div class="communication-mobile-field"><small>Cash FDM</small><div class="communication-mobile-message">${escape(formatCashMoney(record.cashFdm))}</div></div>
      <div class="communication-mobile-field"><small>Card FDM</small><div class="communication-mobile-message">${escape(formatCashMoney(record.cardFdm))}</div></div>
      <div class="communication-mobile-field"><small>Cash (Calc)</small><div class="communication-mobile-message">${escape(formatCashMoney(record.calculatedCash))}</div></div>
      <div class="communication-mobile-field"><small>Dif. Cash</small><div class="communication-mobile-message ${cashDiffValueClass(record.diffCash)}">${escape(formatCashDiffValue(record.diffCash))}</div></div>
      <div class="communication-mobile-field"><small>Dif. Card</small><div class="communication-mobile-message ${cashDiffValueClass(record.diffCard)}">${escape(formatCashDiffValue(record.diffCard))}</div></div>
      <div class="communication-mobile-field communication-mobile-field-full"><small>Justification</small><div class="communication-mobile-message">${escape(record.justification || "-")}</div></div>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" class="ghost${cashItemsAlertClass(record)}" data-cash-action="items-existing" data-id="${escape(record.id)}">${escape(cashItemDiffLabel(record))}</button><button type="button" data-cash-action="edit" data-id="${escape(record.id)}">Edit</button></div></div>`;
  return card;
}

function buildCashInlineCard() {
  const draft = state.cashDraft || emptyCashDraft();
  const computed = cashDraftComputed(draft);
  const card = document.createElement("article");
  card.className = "cash-mobile-card cash-inline-card";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(formatCashDateCompact(draft.day))} Â· ${escape(cashShiftDisplayLabel(draft.shiftName || cashShiftById(draft.shiftId)?.name || ""))}</div>
        <div class="communication-mobile-meta">New shift</div>
      </div>
    </div>
    <div class="communication-mobile-grid">
      <label class="communication-mobile-field communication-mobile-field-full"><small>Name</small><input data-cash-field="name" data-cash-surface="card" data-scope="new" type="text" value="${escape(draft.name)}" /></label>
      <div class="communication-mobile-field"><small>Cash</small><div class="communication-mobile-message"><button type="button" class="ghost" data-cash-action="cash" data-scope="new">${escape(cashSummaryButtonLabel(computed))}</button></div></div>
      <label class="communication-mobile-field"><small>Card POS</small><input class="cash-money-input" data-cash-field="cardPos" data-cash-surface="card" data-scope="new" type="text" inputmode="decimal" value="${escape(String(draft.cardPos ?? ""))}" /></label>
      <label class="communication-mobile-field"><small>Cash FDM</small><input class="cash-money-input" data-cash-field="cashFdm" data-cash-surface="card" data-scope="new" type="text" inputmode="decimal" value="${escape(String(draft.cashFdm ?? ""))}" /></label>
      <label class="communication-mobile-field"><small>Card FDM</small><input class="cash-money-input" data-cash-field="cardFdm" data-cash-surface="card" data-scope="new" type="text" inputmode="decimal" value="${escape(String(draft.cardFdm ?? ""))}" /></label>
      <div class="communication-mobile-field"><small>Cash (Calc)</small><div class="communication-mobile-message">${escape(formatCashMoney(computed.calculatedCash))}</div></div>
      <div class="communication-mobile-field"><small>Dif. Cash</small><div class="communication-mobile-message ${cashDiffValueClass(computed.diffCash)}">${escape(formatCashDiffValue(computed.diffCash))}</div></div>
      <div class="communication-mobile-field"><small>Dif. Card</small><div class="communication-mobile-message ${cashDiffValueClass(computed.diffCard)}">${escape(formatCashDiffValue(computed.diffCard))}</div></div>
      <label class="communication-mobile-field communication-mobile-field-full"><small>Justification</small><input data-cash-field="justification" data-cash-surface="card" data-scope="new" type="text" value="${escape(draft.justification)}" /></label>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" class="ghost${cashItemsAlertClass(computed)}" data-cash-action="items" data-scope="new">${escape(cashItemDiffLabel(computed))}</button><button type="button" data-cash-action="save-new">Add</button></div></div>`;
  return card;
}

function buildCashEditableCard(record, { openMode = false } = {}) {
  const scope = openMode ? "open" : "edit";
  const draft = currentCashDraft(scope, record.id) || record;
  const computed = cashDraftComputed(draft);
  const card = document.createElement("article");
  card.className = "cash-mobile-card cash-inline-card";
  if ((computed.diffCash != null && computed.diffCash !== 0) || computed.diffCard !== 0) card.classList.add("cash-diff-card");
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(formatCashDateCompact(draft.day))} Â· ${escape(cashShiftDisplayLabel(draft.shiftName || ""))}</div>
        <div class="communication-mobile-meta">${escape(draft.name || "-")}</div>
      </div>
    </div>
    <div class="communication-mobile-grid">
      <label class="communication-mobile-field communication-mobile-field-full"><small>Name</small><input data-cash-field="name" data-cash-surface="card" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" value="${escape(draft.name)}" /></label>
      <div class="communication-mobile-field"><small>Cash</small><div class="communication-mobile-message"><button type="button" class="ghost" data-cash-action="cash" data-id="${escape(record.id)}" data-scope="${escape(scope)}">${escape(cashSummaryButtonLabel(computed))}</button></div></div>
      <label class="communication-mobile-field"><small>Card POS</small><input class="cash-money-input" data-cash-field="cardPos" data-cash-surface="card" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" inputmode="decimal" value="${escape(String(draft.cardPos ?? ""))}" /></label>
      <label class="communication-mobile-field"><small>Cash FDM</small><input class="cash-money-input" data-cash-field="cashFdm" data-cash-surface="card" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" inputmode="decimal" value="${escape(String(draft.cashFdm ?? ""))}" /></label>
      <label class="communication-mobile-field"><small>Card FDM</small><input class="cash-money-input" data-cash-field="cardFdm" data-cash-surface="card" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" inputmode="decimal" value="${escape(String(draft.cardFdm ?? ""))}" /></label>
      <div class="communication-mobile-field"><small>Cash (Calc)</small><div class="communication-mobile-message">${escape(formatCashMoney(computed.calculatedCash))}</div></div>
      <div class="communication-mobile-field"><small>Dif. Cash</small><div class="communication-mobile-message ${cashDiffValueClass(computed.diffCash)}">${escape(formatCashDiffValue(computed.diffCash))}</div></div>
      <div class="communication-mobile-field"><small>Dif. Card</small><div class="communication-mobile-message ${cashDiffValueClass(computed.diffCard)}">${escape(formatCashDiffValue(computed.diffCard))}</div></div>
      <label class="communication-mobile-field communication-mobile-field-full"><small>Justification</small><input data-cash-field="justification" data-cash-surface="card" data-scope="${escape(scope)}" data-id="${escape(record.id)}" type="text" value="${escape(draft.justification)}" /></label>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" class="ghost${cashItemsAlertClass(computed)}" data-cash-action="items" data-id="${escape(record.id)}" data-scope="${escape(scope)}">${escape(cashItemDiffLabel(computed))}</button>${openMode
      ? `<button type="button" data-cash-action="save-edit" data-scope="${escape(scope)}" data-id="${escape(record.id)}">Save</button><button type="button" class="ghost" data-cash-action="close-open" data-id="${escape(record.id)}">Close</button>`
      : `<button type="button" data-cash-action="save-edit" data-scope="${escape(scope)}" data-id="${escape(record.id)}">Save</button><button type="button" class="ghost" data-cash-action="cancel-edit" data-id="${escape(record.id)}">Cancel</button>`}</div></div>`;
  return card;
}

function renderCashMobileCards(rows) {
  if (!els.cashMobileCards) return;
  els.cashMobileCards.innerHTML = "";
  const openRecord = getOpenCashRecordClient(state.cashRecords);
  if (!openRecord) {
    els.cashMobileCards.appendChild(buildCashInlineCard());
  }
  if (!rows.length) {
    const empty = document.createElement("div");
    empty.className = "services-mobile-empty";
    empty.textContent = "No cash records found.";
    els.cashMobileCards.appendChild(empty);
    return;
  }
  rows.forEach((record) => {
    const isOpenRow = isCashOpenStatusClient(record.status);
    els.cashMobileCards.appendChild(state.cashEditingId === record.id || isOpenRow
      ? buildCashEditableCard(record, { openMode: isOpenRow })
      : buildCashReadOnlyCard(record));
  });
}

function renderCashWarning() {
  if (!els.cashWarning) return;
  const messages = [];
  const next = getNextExpectedCashRecordClient();
  const shift = cashShiftById(next.shiftId) || { startTime: "00:00", name: next.shiftName };
  const overdue = isCashShiftOverdue(next.day, shift.startTime);
  if (overdue) messages.push(`Missing shift: ${next.day} ${shift.name}`);
  els.cashWarning.hidden = !messages.length;
  els.cashWarning.textContent = messages.join(" | ");
}

function renderCash() {
  if (!canApp("cash")) {
    if (els.cashCount) els.cashCount.textContent = "0 records";
    if (els.cashRows) els.cashRows.innerHTML = '<tr><td colspan="13" class="empty">Your profile has no access to Cash Control.</td></tr>';
    if (els.cashDetailRows) els.cashDetailRows.innerHTML = '<tr><td colspan="19" class="empty">Your profile has no access to Cash Control.</td></tr>';
    if (els.cashItemDetailRows) els.cashItemDetailRows.innerHTML = '<tr><td colspan="4" class="empty">Your profile has no access to Cash Control.</td></tr>';
    if (els.cashResumeRows) els.cashResumeRows.innerHTML = '<tr><td colspan="7" class="empty">Your profile has no access to Cash Control.</td></tr>';
    return;
  }
  const rows = buildComputedCashRowsClient(state.cashRecords, state.cashSettings);
  const openRecord = getOpenCashRecordClient(rows);
  if (openRecord) {
    if (clean(state.cashOpenDraft?.id) !== clean(openRecord.id)) {
      state.cashOpenDraft = clone(openRecord);
    }
  } else {
    state.cashOpenDraft = null;
  }
  if (state.cashEditingId) {
    const editingRecord = rows.find((record) => clean(record.id) === clean(state.cashEditingId));
    if (!editingRecord || isCashOpenStatusClient(editingRecord.status)) {
      state.cashEditingId = "";
      state.cashEditDraft = null;
    } else if (clean(state.cashEditDraft?.id) !== clean(editingRecord.id)) {
      state.cashEditDraft = clone(editingRecord);
    }
  }
  const visibleRows = visibleCashRows(rows, state.cashSettings);
  renderCashScreenTabs();
  const focusTarget = document.activeElement?.matches?.("[data-cash-field]") ? document.activeElement : null;
  const focusField = clean(focusTarget?.dataset?.cashField);
  const focusSurface = clean(focusTarget?.dataset?.cashSurface);
  const focusScope = clean(focusTarget?.dataset?.scope);
  const focusId = clean(focusTarget?.dataset?.id);
  const caretStart = focusTarget && typeof focusTarget.selectionStart === "number" ? focusTarget.selectionStart : null;
  const caretEnd = focusTarget && typeof focusTarget.selectionEnd === "number" ? focusTarget.selectionEnd : null;
  if (els.cashFilterDateFrom) els.cashFilterDateFrom.value = clean(state.cashFilters.dateFrom);
  if (els.cashFilterDateTo) els.cashFilterDateTo.value = clean(state.cashFilters.dateTo);
  if (els.cashFilterShift) {
    const shifts = state.cashSettings?.shifts || DEFAULT_CASH_SETTINGS.shifts;
    const current = clean(state.cashFilters.shift);
    els.cashFilterShift.innerHTML = [`<option value="">All</option>`, ...shifts.map((shift) => `<option value="${escape(clean(shift.id))}">${escape(shift.name)}</option>`)].join("");
    els.cashFilterShift.value = current;
  }
  if (els.cashFilterName) els.cashFilterName.value = clean(state.cashFilters.name);
  if (els.cashCount) els.cashCount.textContent = `${visibleRows.length} record${visibleRows.length === 1 ? "" : "s"}`;
  renderCashWarning();
  renderCashDetailRows(visibleRows);
  renderCashItemDetailRows(visibleRows);
  renderCashResumeRows(visibleRows);
  renderCashMobileCards(visibleRows);
  if (els.cashRows) {
    els.cashRows.innerHTML = "";
    if (!openRecord) {
      els.cashRows.appendChild(buildCashInlineRow());
    }
    visibleRows.forEach((record) => {
      const isOpenRow = isCashOpenStatusClient(record.status);
      els.cashRows.appendChild(state.cashEditingId === record.id || isOpenRow
        ? buildCashEditableRow(record, { openMode: isOpenRow })
        : buildCashReadOnlyRow(record));
    });
  }
  const restoreTarget = focusField
    ? document.querySelector(`[data-cash-field="${CSS.escape(focusField)}"]${focusSurface ? `[data-cash-surface="${CSS.escape(focusSurface)}"]` : ""}[data-scope="${CSS.escape(focusScope || "new")}"]${focusId ? `[data-id="${CSS.escape(focusId)}"]` : ""}`)
    : null;
  if (restoreTarget) {
    restoreTarget.focus();
    if (typeof caretStart === "number" && typeof restoreTarget.setSelectionRange === "function") {
      const forceEnd = ["cardPos", "cashFdm", "cardFdm"].includes(focusField);
      const valueLength = String(restoreTarget.value || "").length;
      const nextStart = forceEnd ? valueLength : Math.min(caretStart, valueLength);
      const nextEnd = forceEnd ? valueLength : Math.min(caretEnd ?? caretStart, valueLength);
      restoreTarget.setSelectionRange(nextStart, nextEnd);
    }
  }
}

function setCashStatus(text) {
  if (els.cashStatus) els.cashStatus.textContent = text;
}

function setCashSettingsTab(tab) {
  state.cashSettingsTab = tab === "min" ? "min" : "config";
  const showConfig = state.cashSettingsTab === "config";
  const showMin = state.cashSettingsTab === "min";
  if (els.cashSettingsConfigTab) {
    els.cashSettingsConfigTab.classList.toggle("active-tab", showConfig);
    els.cashSettingsConfigTab.classList.toggle("ghost", !showConfig);
  }
  if (els.cashSettingsMinTab) {
    els.cashSettingsMinTab.classList.toggle("active-tab", showMin);
    els.cashSettingsMinTab.classList.toggle("ghost", !showMin);
  }
  if (els.cashSettingsConfigPanel) {
    els.cashSettingsConfigPanel.hidden = !showConfig;
    els.cashSettingsConfigPanel.style.display = showConfig ? "" : "none";
  }
  if (els.cashSettingsConfigItemsPanel) {
    els.cashSettingsConfigItemsPanel.hidden = !showConfig;
    els.cashSettingsConfigItemsPanel.style.display = showConfig ? "" : "none";
  }
  if (els.cashSettingsMinPanel) {
    els.cashSettingsMinPanel.hidden = !showMin;
    els.cashSettingsMinPanel.style.display = showMin ? "" : "none";
  }
}

function setCashSettingsStatus(text) {
  if (els.cashSettingsStatus) els.cashSettingsStatus.textContent = text;
}

async function loadCashSettings({ silent = false } = {}) {
  try {
    const result = await api("/api/cash-control-settings");
    state.cashSettings = normalizeCashSettingsClient(result?.settings);
    state.cashSettingsLoaded = true;
    renderCashSettings();
    if (!silent) setCashSettingsStatus("Cash configuration loaded.");
  } catch (e) {
    state.cashSettings = clone(DEFAULT_CASH_SETTINGS);
    renderCashSettings();
    if (!silent) setCashSettingsStatus(`Using default cash configuration (${e.message}).`);
  }
}

async function loadCashData({ silent = false } = {}) {
  try {
    const result = await api("/api/cash-control");
    if (result?.settings) {
      state.cashSettings = normalizeCashSettingsClient(result.settings);
      state.cashSettingsLoaded = true;
    }
    state.cashRecords = (Array.isArray(result?.rows) ? result.rows : []).map((row) => normalizeCashRecordClient(row, state.cashSettings));
    state.cashLoaded = true;
    const openRecord = getOpenCashRecordClient(state.cashRecords);
    state.cashOpenDraft = openRecord ? clone(openRecord) : null;
    if (!state.cashDraft || !clean(state.cashDraft.id)) state.cashDraft = emptyCashDraft();
    renderCash();
    renderCashSettings();
    if (!silent) setCashStatus(`Loaded ${state.cashRecords.length} cash record${state.cashRecords.length === 1 ? "" : "s"}.`);
  } catch (e) {
    state.cashRecords = [];
    renderCash();
    if (!silent) setCashStatus(`Failed to load cash records: ${e.message}`);
  }
}

function renderCashSettings() {
  setCashSettingsTab(state.cashSettingsTab);
  if (els.cashSettingsShiftsBody) {
    els.cashSettingsShiftsBody.innerHTML = (state.cashSettings?.shifts || []).map((shift) => `<tr>
      <td><input data-cash-settings-scope="shift" data-cash-settings-id="${escape(shift.id)}" data-cash-settings-field="name" type="text" value="${escape(shift.name)}" /></td>
      <td><input data-cash-settings-scope="shift" data-cash-settings-id="${escape(shift.id)}" data-cash-settings-field="startTime" type="time" value="${escape(shift.startTime)}" /></td>
      <td><button type="button" class="ghost" data-cash-settings-action="remove-shift" data-id="${escape(shift.id)}">Delete</button></td>
    </tr>`).join("");
  }
  if (els.cashSettingsItemsBody) {
    els.cashSettingsItemsBody.innerHTML = (state.cashSettings?.items || []).map((item) => `<tr>
      <td><input data-cash-settings-scope="item" data-cash-settings-id="${escape(item.id)}" data-cash-settings-field="name" type="text" value="${escape(item.name)}" /></td>
      <td><input data-cash-settings-scope="item" data-cash-settings-id="${escape(item.id)}" data-cash-settings-field="defaultQuantity" type="number" min="0" step="1" value="${escape(String(item.defaultQuantity || 0))}" /></td>
      <td><button type="button" class="ghost" data-cash-settings-action="remove-item" data-id="${escape(item.id)}">Delete</button></td>
    </tr>`).join("");
  }
  if (els.cashSettingsMinBody) {
    const minCash = state.cashSettings?.minCash || DEFAULT_CASH_SETTINGS.minCash;
    const maxCashByDenomination = state.cashSettings?.maxCashByDenomination || DEFAULT_CASH_SETTINGS.maxCashByDenomination;
    els.cashSettingsMinBody.innerHTML = CASH_MIN_ALERT_DENOMINATIONS.map((key) => `<tr>
      <td>${escape(formatCashDenominationLabel(key))}</td>
      <td><input data-cash-settings-scope="min" data-cash-settings-field="${escape(key)}" type="number" min="0" step="1" value="${escape(String(minCash?.[key] || 0))}" /></td>
      <td><input data-cash-settings-scope="max" data-cash-settings-field="${escape(key)}" type="number" min="0" step="1" value="${escape(String(maxCashByDenomination?.[key] || 0))}" /></td>
    </tr>`).join("");
  }
  if (els.cashSettingsManagerAlertEmail) {
    els.cashSettingsManagerAlertEmail.value = (state.cashSettings?.managerAlertEmails || []).join("\n");
  }
  if (els.cashSettingsMinimumEmailEnabled) {
    els.cashSettingsMinimumEmailEnabled.checked = !!state.cashSettings?.minimumCashEmailEnabled;
  }
  if (els.cashSettingsMaximumEmailEnabled) {
    els.cashSettingsMaximumEmailEnabled.checked = !!state.cashSettings?.maximumCashEmailEnabled;
  }
  if (els.cashSettingsMaximumCash) {
    els.cashSettingsMaximumCash.value = state.cashSettings?.maximumCash ? String(state.cashSettings.maximumCash) : "";
  }
}

function onCashSettingsInput(event) {
  const target = event.target;
  if (target === els.cashSettingsManagerAlertEmail) {
    state.cashSettings.managerAlertEmails = parseEmailList(target.value);
    return;
  }
  if (target === els.cashSettingsMinimumEmailEnabled) {
    state.cashSettings.minimumCashEmailEnabled = !!target.checked;
    return;
  }
  if (target === els.cashSettingsMaximumEmailEnabled) {
    state.cashSettings.maximumCashEmailEnabled = !!target.checked;
    return;
  }
  if (target === els.cashSettingsMaximumCash) {
    state.cashSettings.maximumCash = Number((normalizeNumber(target.value) || 0).toFixed(2));
    return;
  }
  const scope = clean(target?.dataset?.cashSettingsScope);
  const id = clean(target?.dataset?.cashSettingsId);
  const field = clean(target?.dataset?.cashSettingsField);
  if (scope === "min" && field) {
    state.cashSettings.minCash[field] = Math.max(0, Math.round(Number(normalizeNumber(target.value) || 0)));
    return;
  }
  if (scope === "max" && field) {
    state.cashSettings.maxCashByDenomination[field] = Math.max(0, Math.round(Number(normalizeNumber(target.value) || 0)));
    return;
  }
  if (!scope || !id || !field) return;
  const list = scope === "shift" ? state.cashSettings.shifts : state.cashSettings.items;
  const row = list.find((item) => clean(item.id) === id);
  if (!row) return;
  row[field] = field === "defaultQuantity" ? Math.max(0, Math.round(Number(normalizeNumber(target.value) || 0))) : target.value;
}

function onCashSettingsAction(event) {
  const button = event.target.closest("[data-cash-settings-action]");
  if (!button) return;
  const action = clean(button.dataset.cashSettingsAction);
  const id = clean(button.dataset.id);
  if (action === "remove-shift") {
    state.cashSettings.shifts = state.cashSettings.shifts.filter((shift) => clean(shift.id) !== id);
  }
  if (action === "remove-item") {
    state.cashSettings.items = state.cashSettings.items.filter((item) => clean(item.id) !== id);
  }
  renderCashSettings();
}

function addCashShiftSetting() {
  const index = (state.cashSettings?.shifts || []).length + 1;
  state.cashSettings.shifts.push({ id: `shift-${index}`, name: `Shift ${index}`, startTime: "00:00" });
  renderCashSettings();
}

function addCashItemSetting() {
  const index = (state.cashSettings?.items || []).length + 1;
  state.cashSettings.items.push({ id: `cash-item-${index}`, name: `Item ${index}`, defaultQuantity: 0 });
  renderCashSettings();
}

async function saveCashSettings() {
  try {
    const result = await api("/api/cash-control-settings", {
      method: "PUT",
      body: { settings: state.cashSettings },
    });
    state.cashSettings = normalizeCashSettingsClient(result?.settings);
    state.cashSettingsLoaded = true;
    state.cashRecords = state.cashRecords.map((record) => normalizeCashRecordClient(record, state.cashSettings));
    renderCashSettings();
    renderCash();
    renderLayout();
    setCashSettingsStatus("Cash configuration saved.");
  } catch (e) {
    setCashSettingsStatus(`Save failed: ${e.message}`);
  }
}

function currentCashDraft(scope = "new", id = "") {
  if (scope === "open") return state.cashOpenDraft;
  if (scope === "edit") return state.cashEditDraft;
  return state.cashDraft;
}

function setCurrentCashDraft(nextDraft, scope = "new") {
  if (scope === "open") state.cashOpenDraft = nextDraft;
  else if (scope === "edit") state.cashEditDraft = nextDraft;
  else state.cashDraft = nextDraft;
}

function onCashTableInput(event) {
  const target = event.target;
  const scope = clean(target?.dataset?.scope || target?.dataset?.cashScope);
  const field = clean(target?.dataset?.cashField);
  const id = clean(target?.dataset?.id);
  if (!field) return;
  const draft = clone(currentCashDraft(scope || "new", id) || emptyCashDraft());
  if (field.startsWith("denom:")) {
    draft.denominations[field.split(":")[1]] = Math.max(0, Math.round(Number(normalizeNumber(target.value) || 0)));
  } else if (field === "cardPos" || field === "cashFdm" || field === "cardFdm") {
    draft[field] = normalizeCashMoneyText(target.value);
  } else {
    draft[field] = target.value;
  }
  setCurrentCashDraft(normalizeCashRecordClient(draft, state.cashSettings), scope || "new");
  renderCash();
}

function onCashFilterInput() {
  state.cashFilters.dateFrom = clean(els.cashFilterDateFrom?.value) || cashDefaultDateFrom();
  state.cashFilters.dateTo = clean(els.cashFilterDateTo?.value);
  state.cashFilters.shift = clean(els.cashFilterShift?.value);
  state.cashFilters.name = clean(els.cashFilterName?.value);
  renderCash();
}

function startCashEdit(id) {
  const record = state.cashRecords.find((row) => clean(row.id) === clean(id));
  if (!record) return;
  state.cashEditingId = record.id;
  state.cashEditDraft = clone(record);
  renderCash();
}

function cancelCashEdit() {
  const editingRecord = state.cashRecords.find((row) => clean(row.id) === clean(state.cashEditingId));
  if (editingRecord && isCashOpenStatusClient(editingRecord.status)) return;
  state.cashEditingId = "";
  state.cashEditDraft = null;
  renderCash();
}

async function saveCashDraft(scope = "new", id = "", { closeRecord = false } = {}) {
  const draft = normalizeCashRecordClient(currentCashDraft(scope, id) || {}, state.cashSettings);
  if (scope === "new") draft.status = "O";
  if (scope === "open") draft.status = "O";
  if (closeRecord) draft.status = "C";
  const error = validateCashDraftClient(draft, { isCreate: scope === "new" });
  if (error) {
    setCashStatus(error);
    showToast(error, "error");
    return;
  }
  try {
    const isUpdate = scope === "edit" || scope === "open";
    const result = await api(isUpdate ? `/api/cash-control?id=${encodeURIComponent(id)}` : "/api/cash-control", {
      method: isUpdate ? "PUT" : "POST",
      body: draft,
    });
    if (result?.settings) {
      state.cashSettings = normalizeCashSettingsClient(result.settings);
      state.cashSettingsLoaded = true;
    }
    state.cashRecords = (Array.isArray(result?.rows) ? result.rows : []).map((row) => normalizeCashRecordClient(row, state.cashSettings));
    state.cashLoaded = true;
    const openRecord = getOpenCashRecordClient(state.cashRecords);
    state.cashOpenDraft = openRecord ? clone(openRecord) : null;
    if (scope === "edit") {
      state.cashEditingId = "";
      state.cashEditDraft = null;
    }
    state.cashDraft = emptyCashDraft();
    renderCash();
    renderLayout();
    const alertEmailError = clean(result?.alertEmailResult?.error);
    if (alertEmailError) {
      setCashStatus(`Saved, but manager alert email failed: ${alertEmailError}`);
      showToast(`Manager alert email failed: ${alertEmailError}`, "error");
      return;
    }
    setCashStatus(closeRecord ? "Cash shift closed." : scope === "edit" ? "Cash record saved." : "Cash record added.");
  } catch (e) {
    setCashStatus(`Save failed: ${e.message}`);
    showToast(`Save failed: ${e.message}`, "error");
  }
}

function renderCashMoneyModal(scope = "new", id = "") {
  const draft = currentCashDraft(scope, id) || emptyCashDraft();
  const computed = cashDraftComputed(draft);
  const focusTarget = document.activeElement?.matches?.("[data-cash-money-key]") ? document.activeElement : null;
  const focusKey = clean(focusTarget?.dataset?.cashMoneyKey);
  const caretStart = focusTarget && typeof focusTarget.selectionStart === "number" ? focusTarget.selectionStart : null;
  const caretEnd = focusTarget && typeof focusTarget.selectionEnd === "number" ? focusTarget.selectionEnd : null;
  if (els.cashMoneyMeta) {
    els.cashMoneyMeta.classList.remove("empty");
    els.cashMoneyMeta.textContent = `${draft.day} · ${draft.shiftName || cashShiftById(draft.shiftId)?.name || ""} · ${draft.name || "-"}`;
  }
  if (els.cashMoneyBody) {
    els.cashMoneyBody.innerHTML = CASH_DENOMINATIONS.map((denom) => {
      const quantity = Math.max(0, Math.round(Number(draft.denominations?.[denom.key] || 0)));
      const value = Number((quantity * denom.value).toFixed(2));
      return `<tr>
        <td>${escape(denom.key)}</td>
        <td><input data-cash-money-key="${escape(denom.key)}" type="text" inputmode="numeric" value="${escape(String(quantity))}" /></td>
        <td>${escape(formatCashMoney(value))}</td>
      </tr>`;
    }).join("");
  }
  if (els.cashMoneyTotal) els.cashMoneyTotal.textContent = `Total cash: ${formatCashMoney(computed.cashTotal)}`;
  if (els.cashMoneyModal) els.cashMoneyModal.hidden = false;
  document.body.classList.add("modal-open");
  if (focusKey) {
    const restoreTarget = document.querySelector(`[data-cash-money-key="${CSS.escape(focusKey)}"]`);
    if (restoreTarget) {
      restoreTarget.focus();
      if (typeof caretStart === "number" && typeof restoreTarget.setSelectionRange === "function") {
        const valueLength = String(restoreTarget.value || "").length;
        const nextStart = valueLength;
        const nextEnd = valueLength;
        restoreTarget.setSelectionRange(nextStart, nextEnd);
      }
    }
  }
}

function openCashMoneyModal(scope = "new", id = "") {
  if (scope === "edit" && id && clean(state.cashEditingId) !== clean(id)) startCashEdit(id);
  state.cashMoneyModalScope = scope || "new";
  state.cashMoneyModalId = id || "";
  state.cashMoneyModalOpen = true;
  renderCashMoneyModal(scope, id);
}

function closeCashMoneyModal() {
  state.cashMoneyModalOpen = false;
  state.cashMoneyModalScope = "new";
  state.cashMoneyModalId = "";
  if (els.cashMoneyModal) els.cashMoneyModal.hidden = true;
  document.body.classList.remove("modal-open");
}

function onCashMoneyModalInput(event) {
  const target = event.target;
  const key = clean(target?.dataset?.cashMoneyKey);
  if (!key) return;
  const scope = clean(state.cashMoneyModalScope) || "new";
  const targetId = clean(state.cashMoneyModalId);
  const draft = clone(currentCashDraft(scope, targetId) || emptyCashDraft());
  draft.denominations[key] = Math.max(0, Math.round(Number(normalizeNumber(target.value) || 0)));
  setCurrentCashDraft(normalizeCashRecordClient(draft, state.cashSettings), scope);
  renderCashMoneyModal(scope, targetId);
}

function saveCashMoneyModal() {
  closeCashMoneyModal();
  renderCash();
}

function ensureCashItemsDraft(scope = "new", id = "") {
  const draft = currentCashDraft(scope, id) || emptyCashDraft();
  state.cashItemsDraft = clone(draft.itemCounts || {});
  state.cashItemsJustificationsDraft = clone(draft.itemJustifications || {});
}

function renderCashItemsModal(scope = "new", id = "") {
  const draft = currentCashDraft(scope, id) || emptyCashDraft();
  const previousRecord = getPreviousCashRecordClient(draft, state.cashRecords, state.cashSettings);
  const computed = cashDraftComputed(draft);
  const focusTarget = document.activeElement?.matches?.("[data-cash-item-field]") ? document.activeElement : null;
  const focusItemId = clean(focusTarget?.dataset?.cashItemId);
  const focusField = clean(focusTarget?.dataset?.cashItemField);
  const caretStart = focusTarget && typeof focusTarget.selectionStart === "number" ? focusTarget.selectionStart : null;
  const caretEnd = focusTarget && typeof focusTarget.selectionEnd === "number" ? focusTarget.selectionEnd : null;
  if (els.cashItemsMeta) {
    els.cashItemsMeta.classList.remove("empty");
    els.cashItemsMeta.textContent = `${draft.day} · ${draft.shiftName || cashShiftById(draft.shiftId)?.name || ""} · ${draft.name || "-"}`;
  }
  if (els.cashItemsBody) {
    els.cashItemsBody.innerHTML = (state.cashSettings?.items || []).map((item) => {
      const counted = state.cashItemsDraft[item.id];
      const diff = counted == null ? null : Number(counted) - Number(item.defaultQuantity || 0);
      const needsJustification = diff != null && diff !== 0;
      const previousJustification = clean(previousRecord?.itemJustifications?.[item.id]);
      return `<tr>
        <td>${escape(item.name)}</td>
        <td>${escape(String(item.defaultQuantity || 0))}</td>
        <td><input data-cash-item-field="count" data-cash-item-id="${escape(item.id)}" type="text" inputmode="numeric" value="${counted == null ? "" : escape(String(counted))}" /></td>
        <td>${escape(diff == null ? "-" : `${diff > 0 ? "+" : ""}${diff}`)}</td>
        <td><input data-cash-item-field="justification" data-cash-item-id="${escape(item.id)}" type="text" value="${escape(state.cashItemsJustificationsDraft[item.id] || "")}" ${needsJustification ? "" : "disabled"} /></td>
        <td>${previousJustification ? `<div class="cash-previous-justification-row"><button type="button" class="ghost cash-copy-previous-button" data-cash-item-action="copy-previous" data-cash-item-id="${escape(item.id)}">Copy</button><div class="cash-previous-justification">${escape(previousJustification)}</div></div>` : "-"}</td>
      </tr>`;
    }).join("");
  }
  if (els.cashItemsStatus) els.cashItemsStatus.textContent = "";
  if (els.cashItemsModal) els.cashItemsModal.hidden = false;
  document.body.classList.add("modal-open");
  const restoreTarget = focusItemId && focusField
    ? document.querySelector(`[data-cash-item-id="${CSS.escape(focusItemId)}"][data-cash-item-field="${CSS.escape(focusField)}"]`)
    : null;
  if (restoreTarget) {
    restoreTarget.focus();
    if (typeof caretStart === "number" && typeof restoreTarget.setSelectionRange === "function") {
      const forceEnd = focusField === "count";
      const valueLength = String(restoreTarget.value || "").length;
      const nextStart = forceEnd ? valueLength : Math.min(caretStart, valueLength);
      const nextEnd = forceEnd ? valueLength : Math.min(caretEnd ?? caretStart, valueLength);
      restoreTarget.setSelectionRange(nextStart, nextEnd);
    }
  }
}

function openCashItemsModal(scope = "new", id = "") {
  if (scope === "edit" && id && clean(state.cashEditingId) !== clean(id)) startCashEdit(id);
  state.cashItemsModalScope = scope || "new";
  state.cashItemsModalId = id || "";
  ensureCashItemsDraft(scope, id);
  state.cashItemsModalOpen = true;
  renderCashItemsModal(scope, id);
}

function closeCashItemsModal() {
  state.cashItemsModalOpen = false;
  state.cashItemsModalScope = "new";
  state.cashItemsModalId = "";
  if (els.cashItemsModal) els.cashItemsModal.hidden = true;
  document.body.classList.remove("modal-open");
}

function onCashItemsModalInput(event) {
  const target = event.target;
  const itemId = clean(target?.dataset?.cashItemId);
  const field = clean(target?.dataset?.cashItemField);
  if (!itemId || !field) return;
  if (field === "count") {
    const numeric = normalizeNumber(target.value);
    state.cashItemsDraft[itemId] = numeric == null ? null : Math.max(0, Math.round(Number(numeric || 0)));
  } else {
    state.cashItemsJustificationsDraft[itemId] = target.value;
  }
  renderCashItemsModal(clean(state.cashItemsModalScope) || "new", clean(state.cashItemsModalId));
}

function onCashItemsModalAction(event) {
  const button = event.target.closest("[data-cash-item-action]");
  if (!button) return;
  const action = clean(button.dataset.cashItemAction);
  const itemId = clean(button.dataset.cashItemId);
  if (action !== "copy-previous" || !itemId) return;
  const scope = clean(state.cashItemsModalScope) || "new";
  const targetId = clean(state.cashItemsModalId);
  const draft = currentCashDraft(scope, targetId) || emptyCashDraft();
  const previousRecord = getPreviousCashRecordClient(draft, state.cashRecords, state.cashSettings);
  const previousJustification = clean(previousRecord?.itemJustifications?.[itemId]);
  if (!previousJustification) return;
  state.cashItemsJustificationsDraft[itemId] = previousJustification;
  renderCashItemsModal(scope, targetId);
  const input = document.querySelector(`[data-cash-item-id="${CSS.escape(itemId)}"][data-cash-item-field="justification"]`);
  if (input) {
    input.focus();
    if (typeof input.setSelectionRange === "function") {
      const length = String(input.value || "").length;
      input.setSelectionRange(length, length);
    }
  }
}

function saveCashItemsModal() {
  const scope = clean(state.cashItemsModalScope) || "new";
  const targetId = clean(state.cashItemsModalId);
  const draft = clone(currentCashDraft(scope, targetId) || emptyCashDraft());
  draft.itemCounts = clone(state.cashItemsDraft);
  draft.itemJustifications = clone(state.cashItemsJustificationsDraft);
  setCurrentCashDraft(normalizeCashRecordClient(draft, state.cashSettings), scope);
  closeCashItemsModal();
  renderCash();
}

function onCashTableAction(event) {
  const button = event.target.closest("[data-cash-action]");
  if (!button) return;
  const action = clean(button.dataset.cashAction);
  const id = clean(button.dataset.id);
  const scope = clean(button.dataset.scope) || (action === "save-edit" || action === "cancel-edit" ? "edit" : "new");
  if (action === "edit" && id) startCashEdit(id);
  if (action === "cancel-edit") cancelCashEdit();
  if (action === "save-new") saveCashDraft("new");
  if (action === "save-edit" && id) saveCashDraft(scope === "open" ? "open" : "edit", id);
  if (action === "close-open" && id) saveCashDraft("open", id, { closeRecord: true });
  if (action === "cash" || action === "cash-existing") openCashMoneyModal(action === "cash-existing" ? "edit" : scope, id);
  if (action === "items" || action === "items-existing") openCashItemsModal(action === "items-existing" ? "edit" : scope, id);
}

function lisbonCurrentTimeClient() {
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/Lisbon",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(new Date());
}

function isCashShiftOverdue(day, startTime) {
  const today = lisbonTodayIsoClient();
  if (!clean(day) || !clean(startTime)) return false;
  if (day < today) return true;
  if (day > today) return false;
  return clean(startTime) <= clean(lisbonCurrentTimeClient());
}

function shouldShowCashAlert() {
  if (!canApp("cash")) return false;
  const next = getNextExpectedCashRecordClient();
  const shift = cashShiftById(next.shiftId) || { startTime: "00:00" };
  return isCashShiftOverdue(next.day, shift.startTime);
}

function normalizeGuestsSettingsClient(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const mappingSource = source.integrationMapping && typeof source.integrationMapping === "object" ? source.integrationMapping : {};
  const sefSource = source.sefCredentials && typeof source.sefCredentials === "object" ? source.sefCredentials : {};
  return {
    sendTime: /^\d{2}:\d{2}$/.test(clean(source.sendTime ?? source.send_time)) ? clean(source.sendTime ?? source.send_time) : DEFAULT_GUESTS_SETTINGS.sendTime,
    integrationMapping: {
      name: clean(mappingSource.name) || DEFAULT_GUESTS_INTEGRATION_MAPPING.name,
      nationality: clean(mappingSource.nationality) || DEFAULT_GUESTS_INTEGRATION_MAPPING.nationality,
      birthDate: clean(mappingSource.birthDate ?? mappingSource.birth_date) || DEFAULT_GUESTS_INTEGRATION_MAPPING.birthDate,
      docNumber: clean(mappingSource.docNumber ?? mappingSource.doc_number) || DEFAULT_GUESTS_INTEGRATION_MAPPING.docNumber,
      docType: clean(mappingSource.docType ?? mappingSource.doc_type) || DEFAULT_GUESTS_INTEGRATION_MAPPING.docType,
      issuerCountry: clean(mappingSource.issuerCountry ?? mappingSource.issuer_country) || DEFAULT_GUESTS_INTEGRATION_MAPPING.issuerCountry,
      residenceCountry: clean(mappingSource.residenceCountry ?? mappingSource.residence_country) || DEFAULT_GUESTS_INTEGRATION_MAPPING.residenceCountry,
      residenceCity: clean(mappingSource.residenceCity ?? mappingSource.residence_city) || DEFAULT_GUESTS_INTEGRATION_MAPPING.residenceCity,
      checkIn: clean(mappingSource.checkIn ?? mappingSource.check_in) || DEFAULT_GUESTS_INTEGRATION_MAPPING.checkIn,
      checkOut: clean(mappingSource.checkOut ?? mappingSource.check_out) || DEFAULT_GUESTS_INTEGRATION_MAPPING.checkOut,
    },
    sefCredentials: {
      unitCode: clean(sefSource.unitCode ?? sefSource.unit_code) || DEFAULT_GUESTS_SETTINGS.sefCredentials.unitCode,
      establishment: clean(sefSource.establishment) || DEFAULT_GUESTS_SETTINGS.sefCredentials.establishment,
      accessKey: clean(sefSource.accessKey ?? sefSource.access_key) || DEFAULT_GUESTS_SETTINGS.sefCredentials.accessKey,
      caCertificate: String(sefSource.caCertificate ?? sefSource.ca_certificate ?? "").trim(),
    },
  };
}

function normalizeGuestApiCallClient(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  return {
    id: clean(source.id),
    createdAt: clean(source.createdAt ?? source.created_at),
    endpoint: clean(source.endpoint),
    requestMethod: clean(source.requestMethod ?? source.request_method) || "POST",
    soapAction: clean(source.soapAction ?? source.soap_action),
    httpStatus: Number.parseInt(source.httpStatus ?? source.http_status, 10) || 0,
    fileNumber: Number.parseInt(source.fileNumber ?? source.file_number, 10) || 0,
    guestCount: Number.parseInt(source.guestCount ?? source.guest_count, 10) || 0,
    success: !!source.success,
    responseMessage: clean(source.responseMessage ?? source.response_message),
    errorMessage: clean(source.errorMessage ?? source.error_message),
    requestDetails: source.requestDetails ?? source.request_details ?? {},
    requestBody: String(source.requestBody ?? source.request_body ?? ""),
    responseBody: String(source.responseBody ?? source.response_body ?? ""),
    responseHeaders: source.responseHeaders ?? source.response_headers ?? {},
  };
}

function emptyGuestDraft() {
  const today = lisbonTodayIsoClient();
  return {
    ha: "H",
    name: "",
    nationality: "",
    nationalityCode: "",
    birthDate: "",
    birthPlace: "",
    docNumber: "",
    docType: "",
    issuerCountry: "",
    issuerCountryCode: "",
    residenceCountry: "",
    residenceCountryCode: "",
    residenceCity: "",
    checkIn: today,
    checkOut: "",
    sentStatus: "pending",
    sentAt: "",
    sendError: "",
    sendBatchNumber: 0,
  };
}

function emptyGuestsBlacklistDraft() {
  return {
    name: "",
    nationality: "",
    nationalityCode: "",
    birthDate: "",
    docNumber: "",
    whatHappened: "",
    occurrenceDate: "",
    whoReported: "",
  };
}

function guestCountryLookupKeyClient(value) {
  return clean(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
}

function resolveGuestCountryClient(value) {
  const raw = clean(value);
  const key = guestCountryLookupKeyClient(raw);
  const entry = (state.guestsCountries || []).find((item) => {
    const code = clean(item.code).toUpperCase();
    const name = clean(item.name);
    const abbr = clean(item.abbr);
    return [code, name, abbr].some((candidate) => guestCountryLookupKeyClient(candidate) === key);
  });
  if (entry) {
    return {
      input: raw || clean(entry.name),
      code: clean(entry.code).toUpperCase(),
      name: clean(entry.name),
      abbr: clean(entry.abbr),
    };
  }
  return {
    input: raw,
    code: /^[A-Z]{3}$/.test(key) ? key : "",
    name: raw,
    abbr: raw,
  };
}

function normalizeGuestDocTypeClient(value) {
  return clean(value).toUpperCase();
}

function coerceGuestDocTypeClient(value) {
  const normalized = normalizeGuestDocTypeClient(value);
  return ["P", "O", "B"].includes(normalized) ? normalized : "O";
}

function normalizeGuestHAClient(value) {
  return clean(value).toUpperCase() === "A" ? "A" : "H";
}

function normalizeGuestDocNumberClient(value) {
  return clean(value).toUpperCase().replace(/\s+/g, "");
}

function normalizeGuestRecordClient(input = {}) {
  const nationality = resolveGuestCountryClient(input.nationality);
  const issuerCountry = resolveGuestCountryClient(input.issuerCountry);
  const residenceCountry = resolveGuestCountryClient(input.residenceCountry);
  return {
    id: clean(input.id),
    ha: normalizeGuestHAClient(input.ha),
    name: clean(input.name),
    nationality: nationality.input,
    nationalityCode: nationality.code,
    birthDate: normalizeDateInput(input.birthDate),
    birthPlace: clean(input.birthPlace),
    docNumber: normalizeGuestDocNumberClient(input.docNumber),
    docType: normalizeGuestDocTypeClient(input.docType),
    issuerCountry: issuerCountry.input,
    issuerCountryCode: issuerCountry.code,
    residenceCountry: residenceCountry.input,
    residenceCountryCode: residenceCountry.code,
    residenceCity: clean(input.residenceCity),
    checkIn: normalizeDateInput(input.checkIn),
    checkOut: normalizeDateInput(input.checkOut),
    sentStatus: ["sent", "error"].includes(clean(input.sentStatus).toLowerCase()) ? clean(input.sentStatus).toLowerCase() : "pending",
    sentAt: clean(input.sentAt),
    sendError: clean(input.sendError),
    sendBatchNumber: Math.max(0, Number.parseInt(input.sendBatchNumber, 10) || 0),
    createdAt: clean(input.createdAt),
    updatedAt: clean(input.updatedAt),
  };
}

function isValidIsoDateClient(value) {
  const raw = clean(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return false;
  const [year, month, day] = raw.split("-").map((part) => Number.parseInt(part, 10));
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function normalizeGuestDateClient(value) {
  const normalized = normalizeDateInput(value);
  return isValidIsoDateClient(normalized) ? normalized : "";
}

function guestIsLocked(record) {
  return clean(record?.sentStatus).toLowerCase() === "sent";
}

function normalizeGuestsBlacklistRecordClient(input = {}) {
  const nationality = resolveGuestCountryClient(input.nationality);
  return {
    id: clean(input.id),
    name: clean(input.name),
    nationality: nationality.input,
    nationalityCode: nationality.code,
    birthDate: normalizeDateInput(input.birthDate),
    docNumber: normalizeGuestDocNumberClient(input.docNumber),
    whatHappened: clean(input.whatHappened),
    occurrenceDate: normalizeDateInput(input.occurrenceDate),
    whoReported: clean(input.whoReported),
    createdAt: clean(input.createdAt),
    updatedAt: clean(input.updatedAt),
  };
}

function sortGuestsRowsClient(rows) {
  return [...(Array.isArray(rows) ? rows : [])].sort((a, b) => {
    const at = new Date(clean(a.createdAt)).getTime() || 0;
    const bt = new Date(clean(b.createdAt)).getTime() || 0;
    return bt - at || clean(a.name).localeCompare(clean(b.name));
  });
}

function sortGuestsBlacklistRowsClient(rows) {
  return [...(Array.isArray(rows) ? rows : [])].sort((a, b) => clean(b.occurrenceDate).localeCompare(clean(a.occurrenceDate)) || clean(a.name).localeCompare(clean(b.name)));
}

function normalizeGuestDescriptionRowClient(input = {}) {
  const colorKey = clean(input.colorKey || input.color_key).toLowerCase();
  return {
    id: clean(input.id),
    room: clean(input.room),
    bed: clean(input.bed).toUpperCase() || "NA",
    guestDescription: clean(input.guestDescription ?? input.guest_description),
    colorKey: Object.prototype.hasOwnProperty.call(GUEST_DESCRIPTION_PALETTES, colorKey) ? colorKey : "blue",
    rowOrder: Number.parseInt(input.rowOrder ?? input.row_order, 10) || 0,
  };
}

function sortGuestDescriptionRowsClient(rows) {
  return [...(Array.isArray(rows) ? rows : [])].sort((a, b) => (Number(a.rowOrder) || 0) - (Number(b.rowOrder) || 0) || clean(a.room).localeCompare(clean(b.room)) || clean(a.bed).localeCompare(clean(b.bed)));
}

function setGuestsStatus(message) {
  if (els.guestsStatus) els.guestsStatus.textContent = message || "";
}

function setGuestsDescriptionsStatus(message) {
  if (els.guestsDescriptionsStatus) els.guestsDescriptionsStatus.textContent = message || "";
}

function setGuestsBlacklistStatus(message) {
  if (els.guestsBlacklistStatus) els.guestsBlacklistStatus.textContent = message || "";
}

function setGuestsSettingsStatus(message) {
  if (els.guestsSettingsStatus) els.guestsSettingsStatus.textContent = message || "";
}

async function loadGuestsSettings({ silent = false } = {}) {
  try {
    const result = await api("/api/guests-settings");
    state.guestsSettings = normalizeGuestsSettingsClient(result?.settings);
    state.guestsCountries = Array.isArray(result?.countries) ? result.countries : state.guestsCountries;
    state.guestsApiCalls = (Array.isArray(result?.apiCalls) ? result.apiCalls : []).map(normalizeGuestApiCallClient);
    state.guestsApiCallsEnabled = result?.apiCallsEnabled !== false;
    state.guestsSettingsLoaded = true;
    renderGuestsCountryOptions();
    renderGuestsSettings();
    if (!silent) setGuestsSettingsStatus("Guests configuration loaded.");
  } catch (e) {
    state.guestsSettings = clone(DEFAULT_GUESTS_SETTINGS);
    state.guestsApiCalls = [];
    state.guestsApiCallsEnabled = false;
    renderGuestsSettings();
    if (!silent) setGuestsSettingsStatus(`Using default guests configuration (${e.message}).`);
  }
}

async function loadGuestsData({ silent = false } = {}) {
  try {
    const [recordsResult, blacklistResult, descriptionsResult] = await Promise.all([
      api("/api/guests"),
      api("/api/guests-blacklist"),
      api("/api/guests-descriptions"),
    ]);
    state.guestsSettings = normalizeGuestsSettingsClient(recordsResult?.settings || state.guestsSettings);
    state.guestsCountries = Array.isArray(recordsResult?.countries) ? recordsResult.countries : Array.isArray(blacklistResult?.countries) ? blacklistResult.countries : state.guestsCountries;
    if (Array.isArray(recordsResult?.apiCalls)) state.guestsApiCalls = recordsResult.apiCalls.map(normalizeGuestApiCallClient);
    if (typeof recordsResult?.apiCallsEnabled === "boolean") state.guestsApiCallsEnabled = recordsResult.apiCallsEnabled;
    state.guestsRows = sortGuestsRowsClient((Array.isArray(recordsResult?.rows) ? recordsResult.rows : []).map(normalizeGuestRecordClient));
    state.guestDescriptionRows = sortGuestDescriptionRowsClient((Array.isArray(descriptionsResult?.rows) ? descriptionsResult.rows : []).map(normalizeGuestDescriptionRowClient));
    state.guestsBlacklist = sortGuestsBlacklistRowsClient((Array.isArray(blacklistResult?.rows) ? blacklistResult.rows : []).map(normalizeGuestsBlacklistRecordClient));
    state.guestsLoaded = true;
    state.guestsSettingsLoaded = true;
    if (!state.guestsEditingId) state.guestsDraft = emptyGuestDraft();
    if (!state.guestsBlacklistEditingId) state.guestsBlacklistDraft = emptyGuestsBlacklistDraft();
    renderGuestsCountryOptions();
    renderGuests();
    renderGuestsSettings();
    if (!silent) {
      setGuestsStatus("Guest records loaded.");
      setGuestsDescriptionsStatus("Guest descriptions loaded.");
      setGuestsBlacklistStatus("Blacklist loaded.");
    }
  } catch (e) {
    state.guestsRows = [];
    state.guestDescriptionRows = [];
    state.guestsBlacklist = [];
    state.guestsSettings = clone(DEFAULT_GUESTS_SETTINGS);
    state.guestsApiCalls = [];
    state.guestsDraft = emptyGuestDraft();
    state.guestsBlacklistDraft = emptyGuestsBlacklistDraft();
    renderGuestsCountryOptions();
    renderGuests();
    renderGuestsSettings();
    if (!silent) {
      setGuestsStatus(`Using default guests data (${e.message}).`);
      setGuestsDescriptionsStatus(`Using default guest descriptions (${e.message}).`);
      setGuestsBlacklistStatus(`Using default blacklist data (${e.message}).`);
    }
  }
}

function onGuestsSettingsInput() {
  const integrationMapping = {};
  GUESTS_INTEGRATION_MAPPING_ROWS.forEach((row) => {
    const select = els.guestsSettingsMappingBody?.querySelector(`[data-guests-mapping-key="${row.key}"]`);
    integrationMapping[row.key] = clean(select?.value) || DEFAULT_GUESTS_INTEGRATION_MAPPING[row.key];
  });
  state.guestsSettings = normalizeGuestsSettingsClient({
    sendTime: els.guestsSettingsSendTime?.value,
    integrationMapping,
    sefCredentials: {
      unitCode: els.guestsSettingsSefUnit?.value,
      establishment: els.guestsSettingsSefEstablishment?.value,
      accessKey: els.guestsSettingsSefAccessKey?.value,
      caCertificate: els.guestsSettingsSefCa?.value,
    },
  });
}

async function saveGuestsSettings() {
  onGuestsSettingsInput();
  try {
    const result = await api("/api/guests-settings", { method: "PUT", body: { settings: state.guestsSettings } });
    state.guestsSettings = normalizeGuestsSettingsClient(result?.settings);
    state.guestsCountries = Array.isArray(result?.countries) ? result.countries : state.guestsCountries;
    state.guestsApiCalls = (Array.isArray(result?.apiCalls) ? result.apiCalls : state.guestsApiCalls).map(normalizeGuestApiCallClient);
    if (typeof result?.apiCallsEnabled === "boolean") state.guestsApiCallsEnabled = result.apiCallsEnabled;
    state.guestsSettingsLoaded = true;
    renderGuestsCountryOptions();
    renderGuestsSettings();
    renderLayout();
    setGuestsSettingsStatus("Guests configuration saved.");
    showToast("Guests configuration saved.", "success");
  } catch (e) {
    setGuestsSettingsStatus(`Save failed: ${e.message}`);
    showToast(`Guests configuration save failed: ${e.message}`, "error");
  }
}

function renderGuestsSettings() {
  renderGuestsSettingsTabs();
  if (els.guestsSettingsSendTime) els.guestsSettingsSendTime.value = clean(state.guestsSettings?.sendTime) || DEFAULT_GUESTS_SETTINGS.sendTime;
  if (els.guestsSettingsMappingBody) {
    els.guestsSettingsMappingBody.innerHTML = GUESTS_INTEGRATION_MAPPING_ROWS.map((row) => {
      const currentValue = clean(state.guestsSettings?.integrationMapping?.[row.key]) || DEFAULT_GUESTS_INTEGRATION_MAPPING[row.key];
      const options = GUESTS_SCREEN_FIELD_OPTIONS.map((option) => `<option value="${escape(option.key)}"${option.key === currentValue ? " selected" : ""}>${escape(option.label)}</option>`).join("");
      return `<tr><td>${escape(row.label)}</td><td><select data-guests-mapping-key="${escape(row.key)}">${options}</select></td></tr>`;
    }).join("");
  }
  if (els.guestsSettingsSefUnit) els.guestsSettingsSefUnit.value = clean(state.guestsSettings?.sefCredentials?.unitCode) || DEFAULT_GUESTS_SETTINGS.sefCredentials.unitCode;
  if (els.guestsSettingsSefEstablishment) els.guestsSettingsSefEstablishment.value = clean(state.guestsSettings?.sefCredentials?.establishment) || DEFAULT_GUESTS_SETTINGS.sefCredentials.establishment;
  if (els.guestsSettingsSefAccessKey) els.guestsSettingsSefAccessKey.value = clean(state.guestsSettings?.sefCredentials?.accessKey) || DEFAULT_GUESTS_SETTINGS.sefCredentials.accessKey;
  if (els.guestsSettingsSefCa) els.guestsSettingsSefCa.value = String(state.guestsSettings?.sefCredentials?.caCertificate || "");
  renderGuestsApiCallsTable();
}

function renderGuestsSettingsTabs() {
  const isConfig = state.guestsSettingsTab === "config";
  const isSef = state.guestsSettingsTab === "sef";
  const isApi = state.guestsSettingsTab === "api";
  if (els.guestsSettingsConfigTab) {
    els.guestsSettingsConfigTab.classList.toggle("active-tab", isConfig);
    els.guestsSettingsConfigTab.classList.toggle("ghost", !isConfig);
  }
  if (els.guestsSettingsSefTab) {
    els.guestsSettingsSefTab.classList.toggle("active-tab", isSef);
    els.guestsSettingsSefTab.classList.toggle("ghost", !isSef);
  }
  if (els.guestsSettingsApiTab) {
    els.guestsSettingsApiTab.classList.toggle("active-tab", isApi);
    els.guestsSettingsApiTab.classList.toggle("ghost", !isApi);
  }
  if (els.guestsSettingsConfigPanel) els.guestsSettingsConfigPanel.hidden = !isConfig;
  if (els.guestsSettingsSefPanel) els.guestsSettingsSefPanel.hidden = !isSef;
  if (els.guestsSettingsApiPanel) els.guestsSettingsApiPanel.hidden = !isApi;
}

function setGuestsSettingsTab(tab) {
  state.guestsSettingsTab = tab === "sef" || tab === "api" ? tab : "config";
  renderGuestsSettingsTabs();
  if (state.guestsSettingsTab === "api" && canSettings("guests")) {
    loadGuestsSettings({ silent: true }).catch(() => {});
  }
}

function renderGuestsApiCallsTable() {
  if (els.guestsSettingsApiNote) {
    els.guestsSettingsApiNote.textContent = state.guestsApiCallsEnabled
      ? "Latest SEF calls saved from the Guests send flow. The request body keeps the activation key masked."
      : "Run the Guest API calls migration in Supabase to enable this technical log table.";
  }
  if (!els.guestsSettingsApiBody) return;
  if (!state.guestsApiCallsEnabled) {
    els.guestsSettingsApiBody.innerHTML = `<tr><td colspan="8" class="muted">API call logging table is not available yet.</td></tr>`;
    return;
  }
  if (!state.guestsApiCalls.length) {
    els.guestsSettingsApiBody.innerHTML = `<tr><td colspan="8" class="muted">No guest API calls recorded yet.</td></tr>`;
    return;
  }
  els.guestsSettingsApiBody.innerHTML = state.guestsApiCalls.map((item) => {
    const message = clean(item.success ? item.responseMessage : (item.errorMessage || item.responseMessage)) || "-";
    const requestText = (() => {
      const details = item.requestDetails && typeof item.requestDetails === "object" ? item.requestDetails : {};
      const detailsText = Object.keys(details).length ? JSON.stringify(details, null, 2) : "";
      return [detailsText, clean(item.requestBody)].filter(Boolean).join("\n\n");
    })();
    const responseHeadersText = item.responseHeaders && typeof item.responseHeaders === "object" && Object.keys(item.responseHeaders).length
      ? JSON.stringify(item.responseHeaders, null, 2)
      : "";
    const responseText = [responseHeadersText, clean(item.responseBody) || clean(item.errorMessage) || "-"].filter(Boolean).join("\n\n");
    return `<tr>
      <td>${escape(formatDateTimeShort(item.createdAt) || "-")}</td>
      <td>${escape(item.success ? "OK" : "Error")}</td>
      <td>${escape(item.httpStatus ? String(item.httpStatus) : "-")}</td>
      <td>${escape(item.guestCount ? String(item.guestCount) : "-")}</td>
      <td class="guest-api-endpoint">${escape(item.endpoint || "-")}</td>
      <td>${escape(message)}</td>
      <td>
        <details class="guest-api-details">
          <summary>View</summary>
          <pre class="guest-api-pre">${escape(requestText || "-")}</pre>
        </details>
      </td>
      <td>
        <details class="guest-api-details">
          <summary>View</summary>
          <pre class="guest-api-pre">${escape(responseText)}</pre>
        </details>
      </td>
    </tr>`;
  }).join("");
}

function renderGuestsCountryOptions() {
  if (!els.guestsCountryList) return;
  const seen = new Set();
  const options = [];
  (state.guestsCountries || []).forEach((country) => {
    const code = clean(country.code).toUpperCase();
    const name = clean(country.name);
    const abbr = clean(country.abbr);
    [
      [name, `${code}${abbr ? ` · ${abbr}` : ""}`],
      [code, name],
      [abbr, `${code} · ${name}`],
    ].forEach(([value, label]) => {
      if (!clean(value)) return;
      const key = `${value}::${label}`;
      if (seen.has(key)) return;
      seen.add(key);
      options.push(`<option value="${escape(value)}" label="${escape(label)}"></option>`);
    });
  });
  els.guestsCountryList.innerHTML = options.join("");
}

function setGuestsScreen(screen) {
  state.guestsScreen = screen === "blacklist" || screen === "descriptions" ? screen : "list";
  renderGuests();
}

function onGuestsFilterInput(event) {
  if (event?.target === els.guestsShowActive) state.guestsFilters.showActive = !!els.guestsShowActive?.checked;
  state.guestsFilters.ha = clean(els.guestsFilterHa?.value);
  state.guestsFilters.search = clean(els.guestsFilterSearch?.value);
  state.guestsFilters.nationality = clean(els.guestsFilterNationality?.value);
  state.guestsFilters.checkInFrom = clean(els.guestsFilterCheckinFrom?.value);
  state.guestsFilters.checkInTo = clean(els.guestsFilterCheckinTo?.value);
  state.guestsFilters.checkOutFrom = clean(els.guestsFilterCheckoutFrom?.value);
  state.guestsFilters.checkOutTo = clean(els.guestsFilterCheckoutTo?.value);
  renderGuests();
}

function guestAgeClient(birthDate, todayIso = lisbonTodayIsoClient()) {
  const birth = clean(birthDate);
  const today = clean(todayIso);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(birth) || !/^\d{4}-\d{2}-\d{2}$/.test(today)) return "";
  let age = Number(today.slice(0, 4)) - Number(birth.slice(0, 4));
  if (today.slice(5) < birth.slice(5)) age -= 1;
  return age >= 0 ? String(age) : "";
}

function normalizeGuestNameMatchClient(value) {
  return clean(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9 ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenizeGuestNameClient(value) {
  return normalizeGuestNameMatchClient(value).split(" ").map((part) => part.trim()).filter((part) => part.length >= 3);
}

function guestBlacklistReasonClient(record, blacklistRecord) {
  const guestDoc = normalizeGuestDocNumberClient(record?.docNumber);
  const blackDoc = normalizeGuestDocNumberClient(blacklistRecord?.docNumber);
  if (guestDoc && blackDoc && guestDoc === blackDoc) return "Doc Number";
  const guestName = normalizeGuestNameMatchClient(record?.name);
  const blackName = normalizeGuestNameMatchClient(blacklistRecord?.name);
  if (guestName && blackName && guestName === blackName) return "Exact Name";
  if (clean(record?.birthDate) && clean(record?.birthDate) === clean(blacklistRecord?.birthDate)) {
    const guestTokens = tokenizeGuestNameClient(record?.name);
    const blackTokens = tokenizeGuestNameClient(blacklistRecord?.name);
    if (guestTokens.some((token) => blackTokens.includes(token))) return "Name + Birth Date";
  }
  return "";
}

function guestBirthdayAlertClient(record, todayIso = lisbonTodayIsoClient()) {
  const birthDate = clean(record?.birthDate);
  const checkIn = clean(record?.checkIn);
  const checkOut = clean(record?.checkOut);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(birthDate) || !/^\d{4}-\d{2}-\d{2}$/.test(checkIn) || !/^\d{4}-\d{2}-\d{2}$/.test(checkOut)) return "";
  const birthdayMd = birthDate.slice(5);
  if (todayIso >= checkIn && todayIso <= checkOut && todayIso.slice(5) === birthdayMd) return "Birthday today";
  const cursor = new Date(`${checkIn}T00:00:00`);
  const limit = new Date(`${checkOut}T00:00:00`);
  let guard = 0;
  while (!Number.isNaN(cursor.getTime()) && !Number.isNaN(limit.getTime()) && cursor <= limit && guard < 400) {
    const dateIso = new Intl.DateTimeFormat("sv-SE", { timeZone: "Europe/Lisbon", year: "numeric", month: "2-digit", day: "2-digit" }).format(cursor);
    if (dateIso.slice(5) === birthdayMd) return "Birthday during stay";
    cursor.setDate(cursor.getDate() + 1);
    guard += 1;
  }
  return "";
}

function shiftGuestIsoClient(isoDate, days) {
  const value = clean(isoDate);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return "";
  const date = new Date(`${value}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return "";
  date.setUTCDate(date.getUTCDate() + Number(days || 0));
  return date.toISOString().slice(0, 10);
}

function isGuestInHouseOnClient(record, dateIso) {
  const target = clean(dateIso);
  const checkIn = clean(record?.checkIn);
  const checkOut = clean(record?.checkOut);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(target) || !/^\d{4}-\d{2}-\d{2}$/.test(checkIn)) return false;
  return checkIn <= target && (!checkOut || checkOut >= target);
}

function getGuestsTopAlertsSummaryText() {
  const rows = getFilteredGuestsRows();
  const todayIso = lisbonTodayIsoClient();
  const tomorrowIso = shiftGuestIsoClient(todayIso, 1);
  const birthdaysToday = rows.filter((row) => isGuestInHouseOnClient(row, todayIso) && guestBirthdayAlertClient(row, todayIso) === "Birthday today").length;
  const birthdaysTomorrow = rows.filter((row) => isGuestInHouseOnClient(row, tomorrowIso) && guestBirthdayAlertClient(row, tomorrowIso) === "Birthday today").length;
  const blacklistsInHouse = rows.filter((row) => isGuestInHouseOnClient(row, todayIso) && guestRowMetaClient(row).isBlacklisted).length;
  return `Birthdays Today: ${birthdaysToday}; Birthdays tomorrow: ${birthdaysTomorrow}; Blacklists inhouse: ${blacklistsInHouse}`;
}

function guestRowMetaClient(record) {
  const blacklistMatch = state.guestsBlacklist
    .map((item) => ({ item, reason: guestBlacklistReasonClient(record, item) }))
    .find((match) => match.reason);
  const birthdayAlert = guestBirthdayAlertClient(record);
  const age = guestAgeClient(record.birthDate);
  const missingCheckout = !clean(record?.checkOut);
  return {
    blacklistMatch,
    birthdayAlert,
    age,
    missingCheckout,
    isBlacklisted: !!blacklistMatch,
  };
}

function guestStatusText(record, meta) {
  const sentStatus = clean(record.sentStatus).toLowerCase();
  const lines = [];
  lines.push(sentStatus === "sent" ? "Sent" : sentStatus === "error" ? "Error" : "Pending");
  if (clean(record.sentAt)) lines.push(`Sent: ${formatDateTimeShort(record.sentAt)}`);
  if (clean(record.sendError)) lines.push(clean(record.sendError));
  return lines;
}

function guestStatusClass(record) {
  const sentStatus = clean(record.sentStatus).toLowerCase();
  if (sentStatus === "sent") return "guest-status-sent";
  if (sentStatus === "error") return "guest-status-error";
  return "guest-status-pending";
}

function guestAlertsMarkup(meta) {
  const alerts = [];
  if (meta?.blacklistMatch) alerts.push('<span class="guest-alert-chip guest-alert-blacklist" title="Blacklist"><span class="guest-alert-icon">\u26d4</span><span class="guest-alert-text">Blacklist</span></span>');
  if (meta?.birthdayAlert) alerts.push(`<span class="guest-alert-chip guest-alert-birthday" title="${escape(meta.birthdayAlert)}"><span class="guest-alert-icon">\ud83c\udf82</span><span class="guest-alert-text">${escape(meta.birthdayAlert)}</span></span>`);
  if (meta?.missingCheckout) alerts.push('<span class="guest-alert-chip guest-alert-missing-co" title="Missing check-out date"><span class="guest-alert-text">missing CO date</span></span>');
  return alerts.length ? `<div class="guest-alerts">${alerts.join("")}</div>` : "";
}

function guestNationalityMatchesFilter(record, filterValue) {
  const filter = clean(filterValue);
  if (!filter) return true;
  const normalizedFilter = guestCountryLookupKeyClient(filter);
  return [record.nationality, record.nationalityCode].some((value) => guestCountryLookupKeyClient(value) === normalizedFilter || guestCountryLookupKeyClient(value).includes(normalizedFilter));
}

function getFilteredGuestsRows() {
  const filters = state.guestsFilters || {};
  const today = lisbonTodayIsoClient();
  const ha = clean(filters.ha).toUpperCase();
  const search = clean(filters.search).toLowerCase();
  const checkInFrom = clean(filters.checkInFrom);
  const checkInTo = clean(filters.checkInTo);
  const checkOutFrom = clean(filters.checkOutFrom);
  const checkOutTo = clean(filters.checkOutTo);
  return sortGuestsRowsClient(state.guestsRows)
    .filter((row) => !filters.showActive || !clean(row.checkOut) || clean(row.checkOut) >= today)
    .filter((row) => !ha || clean(row.ha).toUpperCase() === ha)
    .filter((row) => !search || clean(row.name).toLowerCase().includes(search) || clean(row.docNumber).toLowerCase().includes(search))
    .filter((row) => guestNationalityMatchesFilter(row, filters.nationality))
    .filter((row) => !checkInFrom || clean(row.checkIn) >= checkInFrom)
    .filter((row) => !checkInTo || clean(row.checkIn) <= checkInTo)
    .filter((row) => !checkOutFrom || clean(row.checkOut) >= checkOutFrom)
    .filter((row) => !checkOutTo || clean(row.checkOut) <= checkOutTo);
}

function getFilteredGuestsBlacklistRows() {
  const filters = state.guestsBlacklistFilters || {};
  const search = clean(filters.search).toLowerCase();
  const whoReported = clean(filters.whoReported).toLowerCase();
  return sortGuestsBlacklistRowsClient(state.guestsBlacklist)
    .filter((row) => !search || clean(row.name).toLowerCase().includes(search) || clean(row.docNumber).toLowerCase().includes(search) || clean(row.whatHappened).toLowerCase().includes(search))
    .filter((row) => !whoReported || clean(row.whoReported).toLowerCase().includes(whoReported))
    .filter((row) => guestNationalityMatchesFilter(row, filters.nationality));
}

function buildGuestsExportRows() {
  return getFilteredGuestsRows().map((row) => {
    const meta = guestRowMetaClient(row);
    const alerts = [];
    if (meta?.blacklistMatch) alerts.push("Blacklist");
    if (meta?.birthdayAlert) alerts.push(meta.birthdayAlert);
    if (meta?.missingCheckout) alerts.push("missing CO date");
    return {
      ha: normalizeGuestHAClient(row.ha),
      name: clean(row.name),
      alerts: alerts.join(" | "),
      nationality: clean(row.nationality || row.nationalityCode),
      birthDate: clean(row.birthDate),
      docNumber: clean(row.docNumber),
      docType: clean(row.docType),
      issuerCountry: clean(row.issuerCountry || row.issuerCountryCode),
      checkIn: clean(row.checkIn),
      checkOut: clean(row.checkOut),
      age: meta?.age == null ? "" : String(meta.age),
      status: guestStatusText(row, meta).join(" | "),
    };
  });
}

function exportGuestsToExcel() {
  const rows = buildGuestsExportRows();
  if (!rows.length) {
    showToast("No guest records to export.", "error");
    return;
  }
  const headers = ["HA", "Name", "Alerts", "Nationality", "Birth Date", "Doc. Number", "Doc Type", "Issuer Country", "Check-in", "Check-out", "Age", "Status"];
  const bodyRows = rows.map((row) => [row.ha, row.name, row.alerts, row.nationality, row.birthDate, row.docNumber, row.docType, row.issuerCountry, row.checkIn, row.checkOut, row.age, row.status]);
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><table border="1"><thead><tr>${headers.map((cell) => `<th>${escape(cell)}</th>`).join("")}</tr></thead><tbody>${bodyRows.map((cells) => `<tr>${cells.map((cell) => `<td>${escape(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table></body></html>`;
  downloadBlob(`guests_list_${formatDate(new Date())}.xls`, html, "application/vnd.ms-excel;charset=utf-8;");
  showToast(`Exported ${rows.length} guest record${rows.length === 1 ? "" : "s"} to Excel.`, "success");
}

function guestCanCopyRecord(record, todayIso = lisbonTodayIsoClient()) {
  const checkIn = clean(record?.checkIn);
  return /^\d{4}-\d{2}-\d{2}$/.test(checkIn) && checkIn < clean(todayIso);
}

function buildGuestCopyDraft(record, todayIso = lisbonTodayIsoClient()) {
  const copiedCheckOut = clean(record?.checkOut);
  const nextCheckIn = /^\d{4}-\d{2}-\d{2}$/.test(copiedCheckOut) && copiedCheckOut >= todayIso ? copiedCheckOut : todayIso;
  return {
    ha: normalizeGuestHAClient(record?.ha),
    name: clean(record?.name),
    nationality: clean(record?.nationality || record?.nationalityCode),
    birthDate: clean(record?.birthDate),
    birthPlace: clean(record?.birthPlace),
    docNumber: clean(record?.docNumber),
    docType: clean(record?.docType),
    issuerCountry: clean(record?.issuerCountry || record?.issuerCountryCode),
    residenceCountry: clean(record?.residenceCountry),
    residenceCity: clean(record?.residenceCity),
    checkIn: nextCheckIn,
    checkOut: "",
  };
}

function buildGuestPayload(draft, { isEdit = false } = {}) {
  const payload = {
    ha: normalizeGuestHAClient(draft.ha),
    name: clean(draft.name),
    nationality: clean(draft.nationality),
    birthDate: normalizeGuestDateClient(draft.birthDate),
    birthPlace: clean(draft.birthPlace),
    docNumber: normalizeGuestDocNumberClient(draft.docNumber),
    docType: coerceGuestDocTypeClient(draft.docType),
    issuerCountry: clean(draft.issuerCountry),
    residenceCountry: clean(draft.residenceCountry),
    residenceCity: clean(draft.residenceCity),
    checkIn: normalizeGuestDateClient(draft.checkIn),
    checkOut: normalizeGuestDateClient(draft.checkOut),
  };
  if (isEdit) {
    payload.sentStatus = "pending";
    payload.sentAt = "";
    payload.sendError = "";
    payload.sendBatchNumber = 0;
  }
  return payload;
}

function buildGuestsBlacklistPayload(draft) {
  return {
    name: clean(draft.name),
    nationality: clean(draft.nationality),
    birthDate: normalizeDateInput(draft.birthDate),
    docNumber: normalizeGuestDocNumberClient(draft.docNumber),
    whatHappened: clean(draft.whatHappened),
    occurrenceDate: normalizeDateInput(draft.occurrenceDate),
    whoReported: clean(draft.whoReported),
  };
}

function validateGuestDraftClient(draft) {
  if (!clean(draft?.name)) return "Guest name is required.";
  if (!clean(draft?.birthDate)) return "Birth date is required.";
  if (!normalizeGuestDateClient(draft?.birthDate)) return "Birth date must be a valid date.";
  if (!clean(draft?.docNumber)) return "Document number is required.";
  if (clean(draft?.checkIn) && !normalizeGuestDateClient(draft?.checkIn)) return "Check-in must be a valid date.";
  if (clean(draft?.checkOut) && !normalizeGuestDateClient(draft?.checkOut)) return "Check-out must be a valid date.";
  if (normalizeGuestDateClient(draft?.checkIn) && normalizeGuestDateClient(draft?.checkOut) && normalizeGuestDateClient(draft?.checkOut) < normalizeGuestDateClient(draft?.checkIn)) return "Check-out must be after or equal to check-in.";
  return "";
}

function validateGuestsBlacklistDraftClient(draft) {
  if (!clean(draft?.name) && !clean(draft?.docNumber)) return "Blacklist record requires a name or document number.";
  if (!clean(draft?.occurrenceDate)) return "Occurrence date is required.";
  return "";
}

async function saveGuestRecord(mode = "new", id = "", options = {}) {
  const isEdit = mode === "edit";
  const draft = isEdit ? state.guestsEditDraft : state.guestsDraft;
  const validationError = validateGuestDraftClient(draft);
  if (validationError) {
    setGuestsStatus(validationError);
    showToast(validationError, "error");
    return;
  }
  try {
    const result = await api(isEdit ? `/api/guests?id=${encodeURIComponent(id)}` : "/api/guests", {
      method: isEdit ? "PUT" : "POST",
      body: buildGuestPayload(draft, { isEdit }),
    });
    state.guestsRows = sortGuestsRowsClient((Array.isArray(result?.rows) ? result.rows : []).map(normalizeGuestRecordClient));
    state.guestsSettings = normalizeGuestsSettingsClient(result?.settings || state.guestsSettings);
    state.guestsCountries = Array.isArray(result?.countries) ? result.countries : state.guestsCountries;
    state.guestsLoaded = true;
    state.guestsSettingsLoaded = true;
    state.guestsEditingId = "";
    state.guestsEditDraft = null;
    clearGuestsQuickEdit();
    state.guestsDraft = emptyGuestDraft();
    renderGuestsCountryOptions();
    renderGuests();
    renderGuestsSettings();
    renderLayout();
    setGuestsStatus(isEdit ? "Guest record saved." : "Guest record added.");
    showToast(isEdit ? "Guest record saved." : "Guest record added.", "success");
    if (options?.focusNewRow) {
      requestAnimationFrame(() => {
        state.guestsDraft.issuerCountry = "";
        const issuerInputs = Array.from(document.querySelectorAll('[data-field="issuerCountry"][data-scope="new"]')).filter((element) => element instanceof HTMLElement && element.offsetParent !== null);
        issuerInputs.forEach((input) => {
          input.value = "";
          input.defaultValue = "";
          input.setAttribute("value", "");
        });
        const next = Array.from(document.querySelectorAll('[data-field="name"][data-scope="new"]')).find((element) => element instanceof HTMLElement && element.offsetParent !== null);
        next?.focus();
      });
    }
  } catch (e) {
    setGuestsStatus(`Save failed: ${e.message}`);
    showToast(`Guest save failed: ${e.message}`, "error");
  }
}

async function deleteGuestRecord(id) {
  if (!window.confirm("Delete this guest record?")) return;
  try {
    const result = await api(`/api/guests?id=${encodeURIComponent(id)}`, { method: "DELETE" });
    state.guestsRows = sortGuestsRowsClient((Array.isArray(result?.rows) ? result.rows : []).map(normalizeGuestRecordClient));
    state.guestsEditingId = "";
    state.guestsEditDraft = null;
    clearGuestsQuickEdit();
    renderGuests();
    renderLayout();
    setGuestsStatus("Guest record deleted.");
    showToast("Guest record deleted.", "success");
  } catch (e) {
    setGuestsStatus(`Delete failed: ${e.message}`);
    showToast(`Guest delete failed: ${e.message}`, "error");
  }
}

async function copyGuestRecord(id) {
  const record = state.guestsRows.find((item) => item.id === id);
  if (!record) return;
  try {
    const result = await api("/api/guests", {
      method: "POST",
      body: buildGuestPayload(buildGuestCopyDraft(record)),
    });
    state.guestsRows = sortGuestsRowsClient((Array.isArray(result?.rows) ? result.rows : []).map(normalizeGuestRecordClient));
    state.guestsSettings = normalizeGuestsSettingsClient(result?.settings || state.guestsSettings);
    state.guestsCountries = Array.isArray(result?.countries) ? result.countries : state.guestsCountries;
    state.guestsLoaded = true;
    state.guestsSettingsLoaded = true;
    renderGuestsCountryOptions();
    renderGuests();
    renderGuestsSettings();
    renderLayout();
    setGuestsStatus("Guest record copied.");
    showToast("Guest record copied.", "success");
  } catch (e) {
    setGuestsStatus(`Copy failed: ${e.message}`);
    showToast(`Guest copy failed: ${e.message}`, "error");
  }
}

async function saveGuestsBlacklistRecord(mode = "new", id = "") {
  const isEdit = mode === "edit";
  const draft = isEdit ? state.guestsBlacklistEditDraft : state.guestsBlacklistDraft;
  const validationError = validateGuestsBlacklistDraftClient(draft);
  if (validationError) {
    setGuestsBlacklistStatus(validationError);
    showToast(validationError, "error");
    return;
  }
  try {
    const result = await api(isEdit ? `/api/guests-blacklist?id=${encodeURIComponent(id)}` : "/api/guests-blacklist", {
      method: isEdit ? "PUT" : "POST",
      body: buildGuestsBlacklistPayload(draft),
    });
    state.guestsBlacklist = sortGuestsBlacklistRowsClient((Array.isArray(result?.rows) ? result.rows : []).map(normalizeGuestsBlacklistRecordClient));
    state.guestsCountries = Array.isArray(result?.countries) ? result.countries : state.guestsCountries;
    state.guestsBlacklistEditingId = "";
    state.guestsBlacklistEditDraft = null;
    state.guestsBlacklistDraft = emptyGuestsBlacklistDraft();
    renderGuestsCountryOptions();
    renderGuests();
    renderLayout();
    setGuestsBlacklistStatus(isEdit ? "Blacklist record saved." : "Blacklist record added.");
    showToast(isEdit ? "Blacklist record saved." : "Blacklist record added.", "success");
  } catch (e) {
    setGuestsBlacklistStatus(`Save failed: ${e.message}`);
    showToast(`Blacklist save failed: ${e.message}`, "error");
  }
}

async function deleteGuestsBlacklistRecord(id) {
  if (!window.confirm("Delete this blacklist record?")) return;
  try {
    const result = await api(`/api/guests-blacklist?id=${encodeURIComponent(id)}`, { method: "DELETE" });
    state.guestsBlacklist = sortGuestsBlacklistRowsClient((Array.isArray(result?.rows) ? result.rows : []).map(normalizeGuestsBlacklistRecordClient));
    state.guestsBlacklistEditingId = "";
    state.guestsBlacklistEditDraft = null;
    renderGuests();
    renderLayout();
    setGuestsBlacklistStatus("Blacklist record deleted.");
    showToast("Blacklist record deleted.", "success");
  } catch (e) {
    setGuestsBlacklistStatus(`Delete failed: ${e.message}`);
    showToast(`Blacklist delete failed: ${e.message}`, "error");
  }
}

async function sendPendingGuests() {
  try {
    const result = await api("/api/guests-send", { method: "POST" });
    state.guestsRows = sortGuestsRowsClient((Array.isArray(result?.rows) ? result.rows : []).map(normalizeGuestRecordClient));
    state.guestsSettings = normalizeGuestsSettingsClient(result?.settings || state.guestsSettings);
    await loadGuestsSettings({ silent: true });
    renderGuests();
    renderLayout();
    setGuestsStatus(result?.message || "Guests sent successfully.");
    showToast(result?.message || `Sent ${Number(result?.sent || 0)} guest record${Number(result?.sent || 0) === 1 ? "" : "s"}.`, "success");
  } catch (e) {
    await loadGuestsData({ silent: true });
    await loadGuestsSettings({ silent: true });
    renderLayout();
    setGuestsStatus(`Send failed: ${e.message}`);
    showToast(`Guest send failed: ${e.message}`, "error");
  }
}

function onGuestsDraftInput(event) {
  const field = clean(event.target.dataset.field);
  const scope = clean(event.target.dataset.scope || "new");
  if (!field) return;
  const draft = scope === "edit" ? state.guestsEditDraft : state.guestsDraft;
  if (!draft) return;
  let value = event.target.value;
  if (event.type === "change" && ["birthDate", "checkIn", "checkOut"].includes(field)) {
    const normalized = normalizeGuestDateClient(value);
    if (normalized) value = normalized;
  }
  if (field === "docType") value = clean(value).toUpperCase();
  draft[field] = value;
  event.target.value = value;
}

function guestQuickFieldMarkup(record, field) {
  if (guestIsLocked(record)) {
    return escape(record[field] || "-");
  }
  const isEditing = clean(state.guestsQuickEditId) === clean(record.id) && clean(state.guestsQuickEditField) === clean(field);
  if (!isEditing) {
    const display = field === "ha" ? normalizeGuestHAClient(record.ha) : clean(record[field]) || "-";
    return `<button type="button" class="ghost guest-quick-trigger" data-guests-quick-start="${escape(field)}" data-id="${escape(record.id)}">${escape(display)}</button>`;
  }
  if (field === "ha") {
    return `<select data-guests-quick-input="ha" data-id="${escape(record.id)}"><option value="H" ${normalizeGuestHAClient(record.ha) === "H" ? "selected" : ""}>H</option><option value="A" ${normalizeGuestHAClient(record.ha) === "A" ? "selected" : ""}>A</option></select>`;
  }
  return `<input data-guests-quick-input="${escape(field)}" data-id="${escape(record.id)}" type="date" value="${escape(record[field])}" />`;
}

function startGuestsQuickEdit(id, field) {
  const row = state.guestsRows.find((item) => item.id === id);
  if (!row || guestIsLocked(row)) return;
  state.guestsQuickEditId = id;
  state.guestsQuickEditField = field;
  renderGuests();
  requestAnimationFrame(() => {
    const selector = `[data-guests-quick-input="${CSS.escape(field)}"][data-id="${CSS.escape(id)}"]`;
    const input = document.querySelector(selector);
    if (!(input instanceof HTMLElement)) return;
    input.focus();
    if (typeof input.select === "function" && input.tagName === "INPUT") input.select();
    if (input.tagName === "SELECT" && typeof input.showPicker === "function") {
      try {
        input.showPicker();
      } catch {}
    }
  });
}

function clearGuestsQuickEdit() {
  state.guestsQuickEditId = "";
  state.guestsQuickEditField = "";
}

function onGuestsQuickEditClick(event) {
  const trigger = event.target.closest("[data-guests-quick-start]");
  if (!trigger) return;
  const field = clean(trigger.dataset.guestsQuickStart);
  const id = clean(trigger.dataset.id);
  if (!field || !id) return;
  startGuestsQuickEdit(id, field);
}

async function saveGuestQuickEdit(id, field, rawValue) {
  const row = state.guestsRows.find((item) => item.id === id);
  if (!row) return "";
  if (guestIsLocked(row)) {
    throw new Error("Sent guest records cannot be modified.");
  }
  let value = rawValue;
  if (field === "ha") value = normalizeGuestHAClient(rawValue);
  if (field === "checkIn" || field === "checkOut") {
    value = clean(rawValue) ? normalizeGuestDateClient(rawValue) : "";
    if (clean(rawValue) && !value) {
      throw new Error(`${field === "checkIn" ? "Check-in" : "Check-out"} must be a valid date.`);
    }
  }
  const draft = {
    ...row,
    [field]: value,
  };
  const validationError = validateGuestDraftClient(draft);
  if (validationError) throw new Error(validationError);
  const result = await api(`/api/guests?id=${encodeURIComponent(id)}`, {
    method: "PUT",
    body: buildGuestPayload(draft),
  });
  state.guestsRows = sortGuestsRowsClient((Array.isArray(result?.rows) ? result.rows : []).map(normalizeGuestRecordClient));
  state.guestsSettings = normalizeGuestsSettingsClient(result?.settings || state.guestsSettings);
  state.guestsCountries = Array.isArray(result?.countries) ? result.countries : state.guestsCountries;
  state.guestsLoaded = true;
  state.guestsSettingsLoaded = true;
  renderGuestsCountryOptions();
  renderGuests();
  renderGuestsSettings();
  renderLayout();
  return field === "ha" ? value : value || "";
}

async function onGuestsQuickEditChange(event) {
  const field = clean(event.target?.dataset?.guestsQuickInput);
  const id = clean(event.target?.dataset?.id);
  if (!field || !id) return;
  try {
    const value = await saveGuestQuickEdit(id, field, event.target.value);
    clearGuestsQuickEdit();
    renderGuests();
    setGuestsStatus("Guest record saved.");
    event.target.value = value;
  } catch (e) {
    const current = state.guestsRows.find((item) => item.id === id);
    event.target.value = field === "ha" ? normalizeGuestHAClient(current?.[field]) : clean(current?.[field]);
    clearGuestsQuickEdit();
    renderGuests();
    setGuestsStatus(`Save failed: ${e.message}`);
    showToast(`Guest save failed: ${e.message}`, "error");
  }
}

function onGuestsBlacklistDraftInput(event) {
  const field = clean(event.target.dataset.field);
  const scope = clean(event.target.dataset.scope || "new");
  if (!field) return;
  const draft = scope === "edit" ? state.guestsBlacklistEditDraft : state.guestsBlacklistDraft;
  if (!draft) return;
  draft[field] = event.target.value;
  if (event.type === "change") renderGuests();
}

function onGuestsBlacklistFilterInput() {
  state.guestsBlacklistFilters.search = clean(els.guestsBlacklistFilterSearch?.value);
  state.guestsBlacklistFilters.whoReported = clean(els.guestsBlacklistFilterReported?.value);
  state.guestsBlacklistFilters.nationality = clean(els.guestsBlacklistFilterNationality?.value);
  renderGuests();
}

async function onGuestsAction(event) {
  const button = event.target.closest("button[data-guests-action]");
  if (!button) return;
  const action = clean(button.dataset.guestsAction);
  const id = clean(button.dataset.id);
  if (action === "save-inline") {
    await saveGuestRecord("new");
    return;
  }
  if (action === "edit" && id) {
    const row = state.guestsRows.find((item) => item.id === id);
    if (!row) return;
    if (guestIsLocked(row)) {
      showToast("Sent guest records cannot be modified.", "error");
      return;
    }
    state.guestsEditingId = id;
    clearGuestsQuickEdit();
    state.guestsEditDraft = { ...row };
    renderGuests();
    return;
  }
  if (action === "save-edit" && id) {
    await saveGuestRecord("edit", id);
    return;
  }
  if (action === "cancel-edit") {
    state.guestsEditingId = "";
    state.guestsEditDraft = null;
    clearGuestsQuickEdit();
    renderGuests();
    return;
  }
  if (action === "delete" && id) {
    const row = state.guestsRows.find((item) => item.id === id);
    if (guestIsLocked(row)) {
      showToast("Sent guest records cannot be modified.", "error");
      return;
    }
    await deleteGuestRecord(id);
    return;
  }
  if (action === "copy" && id) {
    await copyGuestRecord(id);
  }
}

async function onGuestsBlacklistAction(event) {
  const button = event.target.closest("button[data-guests-blacklist-action]");
  if (!button) return;
  const action = clean(button.dataset.guestsBlacklistAction);
  const id = clean(button.dataset.id);
  if (action === "save-inline") {
    await saveGuestsBlacklistRecord("new");
    return;
  }
  if (action === "edit" && id) {
    const row = state.guestsBlacklist.find((item) => item.id === id);
    if (!row) return;
    state.guestsBlacklistEditingId = id;
    state.guestsBlacklistEditDraft = { ...row };
    renderGuests();
    return;
  }
  if (action === "save-edit" && id) {
    await saveGuestsBlacklistRecord("edit", id);
    return;
  }
  if (action === "cancel-edit") {
    state.guestsBlacklistEditingId = "";
    state.guestsBlacklistEditDraft = null;
    renderGuests();
    return;
  }
  if (action === "delete" && id) {
    await deleteGuestsBlacklistRecord(id);
  }
}

async function onGuestsKeydown(event) {
  const field = clean(event.target?.dataset?.field);
  const scope = clean(event.target?.dataset?.scope || "new");
  const orderedFields = ["name", "nationality", "birthDate", "docNumber", "docType", "issuerCountry"];
  if (scope === "new" && field && orderedFields.includes(field) && event.key === "Tab") {
    const index = orderedFields.indexOf(field);
    if (event.shiftKey) {
      if (index > 0) {
        event.preventDefault();
        const previous = Array.from(document.querySelectorAll(`[data-scope="new"][data-field="${orderedFields[index - 1]}"]`)).find((element) => element instanceof HTMLElement && element.offsetParent !== null);
        previous?.focus();
      }
      return;
    }
    event.preventDefault();
    if (field === "issuerCountry") {
      await saveGuestRecord("new", "", { focusNewRow: true });
      return;
    }
    const next = Array.from(document.querySelectorAll(`[data-scope="new"][data-field="${orderedFields[index + 1]}"]`)).find((element) => element instanceof HTMLElement && element.offsetParent !== null);
    next?.focus();
    return;
  }
  if (field === "issuerCountry" && scope === "new" && event.key === "Tab" && !event.shiftKey) {
    event.preventDefault();
    await saveGuestRecord("new", "", { focusNewRow: true });
  }
}

function guestRowBackground(record, meta) {
  return meta?.isBlacklisted ? hexToRgba("#b12030", 0.5) : "#ffffff";
}

function guestStatusMarkup(record, meta) {
  return `<div class="guest-status-stack ${guestStatusClass(record)}">${guestStatusText(record, meta).map((line, index) => `<span${index > 0 ? ' class="guest-status-note"' : ""}>${escape(line)}</span>`).join("")}</div>`;
}

function buildGuestsInlineRow() {
  const draft = state.guestsDraft || emptyGuestDraft();
  const tr = document.createElement("tr");
  tr.className = "inline-editor sticky-new-row";
  tr.innerHTML = `<td><select data-field="ha" data-scope="new" tabindex="-1"><option value="H" ${draft.ha === "H" ? "selected" : ""}>H</option><option value="A" ${draft.ha === "A" ? "selected" : ""}>A</option></select></td>
    <td><input data-field="name" data-scope="new" value="${escape(draft.name)}" /></td>
    <td><input data-field="nationality" data-scope="new" list="guests-country-list" value="${escape(draft.nationality)}" /></td>
    <td><input data-field="birthDate" data-scope="new" type="text" value="${escape(draft.birthDate)}" /></td>
    <td><input data-field="docNumber" data-scope="new" value="${escape(draft.docNumber)}" /></td>
    <td><input data-field="docType" data-scope="new" type="text" value="${escape(draft.docType)}" /></td>
    <td><input data-field="issuerCountry" data-scope="new" list="guests-country-list" autocomplete="new-password" readonly value="${escape(draft.issuerCountry)}" /></td>
    <td><input data-field="checkIn" data-scope="new" type="date" tabindex="-1" value="${escape(draft.checkIn)}" /></td>
    <td><input data-field="checkOut" data-scope="new" type="date" tabindex="-1" value="${escape(draft.checkOut)}" /></td>
    <td>${escape(guestAgeClient(draft.birthDate))}</td>
    <td class="row-actions"><button type="button" data-guests-action="save-inline">Add</button></td>
    <td>${guestStatusMarkup({ sentStatus: "pending", sendError: "" }, { blacklistMatch: null, birthdayAlert: "" })}</td>`;
  tr.style.backgroundColor = "#ffffff";
  return tr;
}

function buildGuestsReadOnlyRow(record) {
  const meta = guestRowMetaClient(record);
  const locked = guestIsLocked(record);
  const canCopy = guestCanCopyRecord(record);
  const actions = [
    canCopy ? `<button type="button" class="ghost" data-guests-action="copy" data-id="${escape(record.id)}">Copy</button>` : "",
    !locked ? `<button type="button" class="ghost" data-guests-action="edit" data-id="${escape(record.id)}">Edit</button>` : "",
    !locked ? `<button type="button" class="danger" data-guests-action="delete" data-id="${escape(record.id)}">Delete</button>` : "",
  ].filter(Boolean).join("");
  const tr = document.createElement("tr");
  const nameCell = `<div class="guest-name-stack"><div class="guest-name-label">${escape(record.name)}</div>${guestAlertsMarkup(meta)}</div>`;
  tr.innerHTML = `<td>${guestQuickFieldMarkup(record, "ha")}</td>
    <td>${nameCell}</td>
    <td>${escape(record.nationality || record.nationalityCode || "-")}</td>
    <td>${escape(record.birthDate || "-")}</td>
    <td>${escape(record.docNumber || "-")}</td>
    <td>${escape(record.docType || "-")}</td>
    <td>${escape(record.issuerCountry || record.issuerCountryCode || "-")}</td>
    <td>${guestQuickFieldMarkup(record, "checkIn")}</td>
    <td>${guestQuickFieldMarkup(record, "checkOut")}</td>
    <td>${escape(meta.age || "-")}</td>
    <td class="row-actions">${actions || "-"}</td>
    <td>${guestStatusMarkup(record, meta)}</td>`;
  tr.style.backgroundColor = guestRowBackground(record, meta);
  return tr;
}

function buildGuestsEditableRow(record) {
  const draft = state.guestsEditDraft || record;
  const meta = guestRowMetaClient(draft);
  const tr = document.createElement("tr");
  tr.className = "inline-editor";
  const nameCell = `<div class="guest-name-stack"><input data-field="name" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.name)}" />${guestAlertsMarkup(meta)}</div>`;
  tr.innerHTML = `<td><select data-field="ha" data-scope="edit" data-id="${escape(record.id)}"><option value="H" ${draft.ha === "H" ? "selected" : ""}>H</option><option value="A" ${draft.ha === "A" ? "selected" : ""}>A</option></select></td>
    <td>${nameCell}</td>
    <td><input data-field="nationality" data-scope="edit" data-id="${escape(record.id)}" list="guests-country-list" value="${escape(draft.nationality)}" /></td>
    <td><input data-field="birthDate" data-scope="edit" data-id="${escape(record.id)}" type="text" value="${escape(draft.birthDate)}" /></td>
    <td><input data-field="docNumber" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.docNumber)}" /></td>
    <td><input data-field="docType" data-scope="edit" data-id="${escape(record.id)}" type="text" value="${escape(draft.docType)}" /></td>
    <td><input data-field="issuerCountry" data-scope="edit" data-id="${escape(record.id)}" list="guests-country-list" value="${escape(draft.issuerCountry)}" /></td>
    <td><input data-field="checkIn" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.checkIn)}" /></td>
    <td><input data-field="checkOut" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.checkOut)}" /></td>
    <td>${escape(guestAgeClient(draft.birthDate) || "-")}</td>
    <td class="row-actions"><button type="button" data-guests-action="save-edit" data-id="${escape(record.id)}">Save</button><button type="button" class="ghost" data-guests-action="cancel-edit" data-id="${escape(record.id)}">Cancel</button></td>
    <td>${guestStatusMarkup(draft, meta)}</td>`;
  tr.style.backgroundColor = guestRowBackground(draft, meta);
  return tr;
}

function buildGuestsInlineCard() {
  const draft = state.guestsDraft || emptyGuestDraft();
  const card = document.createElement("article");
  card.className = "guests-mobile-card";
  card.innerHTML = `<div class="communication-mobile-grid">
      <label class="communication-mobile-field"><small>HA</small><select data-field="ha" data-scope="new" tabindex="-1"><option value="H" ${draft.ha === "H" ? "selected" : ""}>H</option><option value="A" ${draft.ha === "A" ? "selected" : ""}>A</option></select></label>
      <label class="communication-mobile-field communication-mobile-field-full"><small>Name</small><input data-field="name" data-scope="new" value="${escape(draft.name)}" /></label>
      <label class="communication-mobile-field"><small>Nationality</small><input data-field="nationality" data-scope="new" list="guests-country-list" value="${escape(draft.nationality)}" /></label>
      <label class="communication-mobile-field"><small>Birth Date</small><input data-field="birthDate" data-scope="new" type="text" value="${escape(draft.birthDate)}" /></label>
      <label class="communication-mobile-field"><small>Doc. Number</small><input data-field="docNumber" data-scope="new" value="${escape(draft.docNumber)}" /></label>
      <label class="communication-mobile-field"><small>Doc Type</small><input data-field="docType" data-scope="new" type="text" value="${escape(draft.docType)}" /></label>
      <label class="communication-mobile-field"><small>Issuer Country</small><input data-field="issuerCountry" data-scope="new" list="guests-country-list" autocomplete="new-password" readonly value="${escape(draft.issuerCountry)}" /></label>
      <label class="communication-mobile-field"><small>Check-in</small><input data-field="checkIn" data-scope="new" type="date" tabindex="-1" value="${escape(draft.checkIn)}" /></label>
      <label class="communication-mobile-field"><small>Check-out</small><input data-field="checkOut" data-scope="new" type="date" tabindex="-1" value="${escape(draft.checkOut)}" /></label>
      <div class="communication-mobile-field"><small>Age</small><div class="communication-mobile-message">${escape(guestAgeClient(draft.birthDate) || "-")}</div></div>
      <div class="communication-mobile-field communication-mobile-field-full"><small>Status</small>${guestStatusMarkup({ sentStatus: "pending", sendError: "" }, { blacklistMatch: null, birthdayAlert: "" })}</div>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" data-guests-action="save-inline">Add</button></div></div>`;
  return card;
}

function buildGuestsReadOnlyCard(record) {
  const meta = guestRowMetaClient(record);
  const locked = guestIsLocked(record);
  const canCopy = guestCanCopyRecord(record);
  const actions = [
    canCopy ? `<button type="button" class="ghost" data-guests-action="copy" data-id="${escape(record.id)}">Copy</button>` : "",
    !locked ? `<button type="button" class="ghost" data-guests-action="edit" data-id="${escape(record.id)}">Edit</button>` : "",
    !locked ? `<button type="button" class="danger" data-guests-action="delete" data-id="${escape(record.id)}">Delete</button>` : "",
  ].filter(Boolean).join("");
  const card = document.createElement("article");
  card.className = "guests-mobile-card";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="guest-name-stack"><div class="service-mobile-request">${escape(record.name)}</div>${guestAlertsMarkup(meta)}</div>
        <div class="communication-mobile-meta">${escape(record.checkIn || "-")} · ${escape(record.ha)}</div>
      </div>
      <div class="group-mobile-total"><strong>${escape(meta.age || "-")}</strong><small>Age</small></div>
    </div>
    <div class="communication-mobile-grid">
      <div class="communication-mobile-field"><small>Doc</small><div class="communication-mobile-message">${escape(record.docNumber || "-")}</div></div>
      <div class="communication-mobile-field"><small>HA</small><div class="communication-mobile-message">${guestQuickFieldMarkup(record, "ha")}</div></div>
      <div class="communication-mobile-field"><small>Nationality</small><div class="communication-mobile-message">${escape(record.nationality || "-")}</div></div>
      <div class="communication-mobile-field"><small>Birth Date</small><div class="communication-mobile-message">${escape(record.birthDate || "-")}</div></div>
      <div class="communication-mobile-field"><small>Doc Type</small><div class="communication-mobile-message">${escape(record.docType || "-")}</div></div>
      <div class="communication-mobile-field"><small>Issuer</small><div class="communication-mobile-message">${escape(record.issuerCountry || "-")}</div></div>
      <div class="communication-mobile-field"><small>Check-in</small><div class="communication-mobile-message">${guestQuickFieldMarkup(record, "checkIn")}</div></div>
      <div class="communication-mobile-field"><small>Check-out</small><div class="communication-mobile-message">${guestQuickFieldMarkup(record, "checkOut")}</div></div>
      <div class="communication-mobile-field communication-mobile-field-full"><small>Status</small>${guestStatusMarkup(record, meta)}</div>
    </div>
    <div class="communication-mobile-footer">${actions ? `<div class="row-actions">${actions}</div>` : ""}</div>`;
  card.style.backgroundColor = guestRowBackground(record, meta);
  return card;
}

function buildGuestsEditableCard(record) {
  const draft = state.guestsEditDraft || record;
  const meta = guestRowMetaClient(draft);
  const card = document.createElement("article");
  card.className = "guests-mobile-card";
  card.innerHTML = `<div class="communication-mobile-grid">
      <label class="communication-mobile-field"><small>HA</small><select data-field="ha" data-scope="edit" data-id="${escape(record.id)}"><option value="H" ${draft.ha === "H" ? "selected" : ""}>H</option><option value="A" ${draft.ha === "A" ? "selected" : ""}>A</option></select></label>
      <label class="communication-mobile-field communication-mobile-field-full"><small>Name</small><div class="guest-name-stack"><input data-field="name" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.name)}" />${guestAlertsMarkup(meta)}</div></label>
      <label class="communication-mobile-field"><small>Nationality</small><input data-field="nationality" data-scope="edit" data-id="${escape(record.id)}" list="guests-country-list" value="${escape(draft.nationality)}" /></label>
      <label class="communication-mobile-field"><small>Birth Date</small><input data-field="birthDate" data-scope="edit" data-id="${escape(record.id)}" type="text" value="${escape(draft.birthDate)}" /></label>
      <label class="communication-mobile-field"><small>Doc. Number</small><input data-field="docNumber" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.docNumber)}" /></label>
      <label class="communication-mobile-field"><small>Doc Type</small><input data-field="docType" data-scope="edit" data-id="${escape(record.id)}" type="text" value="${escape(draft.docType)}" /></label>
      <label class="communication-mobile-field"><small>Issuer Country</small><input data-field="issuerCountry" data-scope="edit" data-id="${escape(record.id)}" list="guests-country-list" value="${escape(draft.issuerCountry)}" /></label>
      <label class="communication-mobile-field"><small>Check-in</small><input data-field="checkIn" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.checkIn)}" /></label>
      <label class="communication-mobile-field"><small>Check-out</small><input data-field="checkOut" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.checkOut)}" /></label>
      <div class="communication-mobile-field"><small>Age</small><div class="communication-mobile-message">${escape(guestAgeClient(draft.birthDate) || "-")}</div></div>
      <div class="communication-mobile-field communication-mobile-field-full"><small>Status</small>${guestStatusMarkup(draft, meta)}</div>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" data-guests-action="save-edit" data-id="${escape(record.id)}">Save</button><button type="button" class="ghost" data-guests-action="cancel-edit" data-id="${escape(record.id)}">Cancel</button></div></div>`;
  card.style.backgroundColor = guestRowBackground(draft, meta);
  return card;
}

function buildGuestsBlacklistInlineRow() {
  const draft = state.guestsBlacklistDraft || emptyGuestsBlacklistDraft();
  const tr = document.createElement("tr");
  tr.className = "inline-editor sticky-new-row";
  tr.innerHTML = `<td><input data-field="name" data-scope="new" value="${escape(draft.name)}" /></td>
    <td><input data-field="nationality" data-scope="new" list="guests-country-list" value="${escape(draft.nationality)}" /></td>
    <td><input data-field="birthDate" data-scope="new" type="date" value="${escape(draft.birthDate)}" /></td>
    <td><input data-field="docNumber" data-scope="new" value="${escape(draft.docNumber)}" /></td>
    <td><input data-field="whatHappened" data-scope="new" value="${escape(draft.whatHappened)}" /></td>
    <td><input data-field="occurrenceDate" data-scope="new" type="date" value="${escape(draft.occurrenceDate)}" /></td>
    <td><input data-field="whoReported" data-scope="new" value="${escape(draft.whoReported)}" /></td>
    <td class="row-actions"><button type="button" data-guests-blacklist-action="save-inline">Add</button></td>`;
  return tr;
}

function buildGuestsBlacklistReadOnlyRow(record) {
  const tr = document.createElement("tr");
  tr.innerHTML = `<td>${escape(record.name || "-")}</td>
    <td>${escape(record.nationality || "-")}</td>
    <td>${escape(record.birthDate || "-")}</td>
    <td>${escape(record.docNumber || "-")}</td>
    <td>${escape(record.whatHappened || "-")}</td>
    <td>${escape(record.occurrenceDate || "-")}</td>
    <td>${escape(record.whoReported || "-")}</td>
    <td class="row-actions"><button type="button" class="ghost" data-guests-blacklist-action="edit" data-id="${escape(record.id)}">Edit</button><button type="button" class="danger" data-guests-blacklist-action="delete" data-id="${escape(record.id)}">Delete</button></td>`;
  return tr;
}

function buildGuestsBlacklistEditableRow(record) {
  const draft = state.guestsBlacklistEditDraft || record;
  const tr = document.createElement("tr");
  tr.className = "inline-editor";
  tr.innerHTML = `<td><input data-field="name" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.name)}" /></td>
    <td><input data-field="nationality" data-scope="edit" data-id="${escape(record.id)}" list="guests-country-list" value="${escape(draft.nationality)}" /></td>
    <td><input data-field="birthDate" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.birthDate)}" /></td>
    <td><input data-field="docNumber" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.docNumber)}" /></td>
    <td><input data-field="whatHappened" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.whatHappened)}" /></td>
    <td><input data-field="occurrenceDate" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.occurrenceDate)}" /></td>
    <td><input data-field="whoReported" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.whoReported)}" /></td>
    <td class="row-actions"><button type="button" data-guests-blacklist-action="save-edit" data-id="${escape(record.id)}">Save</button><button type="button" class="ghost" data-guests-blacklist-action="cancel-edit" data-id="${escape(record.id)}">Cancel</button></td>`;
  return tr;
}

function buildGuestsBlacklistInlineCard() {
  const draft = state.guestsBlacklistDraft || emptyGuestsBlacklistDraft();
  const card = document.createElement("article");
  card.className = "guests-mobile-card";
  card.innerHTML = `<div class="communication-mobile-grid">
      <label class="communication-mobile-field communication-mobile-field-full"><small>Name</small><input data-field="name" data-scope="new" value="${escape(draft.name)}" /></label>
      <label class="communication-mobile-field"><small>Nationality</small><input data-field="nationality" data-scope="new" list="guests-country-list" value="${escape(draft.nationality)}" /></label>
      <label class="communication-mobile-field"><small>Birth Date</small><input data-field="birthDate" data-scope="new" type="date" value="${escape(draft.birthDate)}" /></label>
      <label class="communication-mobile-field"><small>Doc. Number</small><input data-field="docNumber" data-scope="new" value="${escape(draft.docNumber)}" /></label>
      <label class="communication-mobile-field"><small>Occurrence Date</small><input data-field="occurrenceDate" data-scope="new" type="date" value="${escape(draft.occurrenceDate)}" /></label>
      <label class="communication-mobile-field communication-mobile-field-full"><small>Who Reported</small><input data-field="whoReported" data-scope="new" value="${escape(draft.whoReported)}" /></label>
      <label class="communication-mobile-field communication-mobile-field-full"><small>What Happened?</small><textarea data-field="whatHappened" data-scope="new" rows="3">${escape(draft.whatHappened)}</textarea></label>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" data-guests-blacklist-action="save-inline">Add</button></div></div>`;
  return card;
}

function buildGuestsBlacklistReadOnlyCard(record) {
  const card = document.createElement("article");
  card.className = "guests-mobile-card";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(record.name || "-")}</div>
        <div class="communication-mobile-meta">${escape(record.occurrenceDate || "-")}</div>
      </div>
    </div>
    <div class="communication-mobile-grid">
      <div class="communication-mobile-field"><small>Nationality</small><div class="communication-mobile-message">${escape(record.nationality || "-")}</div></div>
      <div class="communication-mobile-field"><small>Birth Date</small><div class="communication-mobile-message">${escape(record.birthDate || "-")}</div></div>
      <div class="communication-mobile-field"><small>Doc. Number</small><div class="communication-mobile-message">${escape(record.docNumber || "-")}</div></div>
      <div class="communication-mobile-field"><small>Who Reported</small><div class="communication-mobile-message">${escape(record.whoReported || "-")}</div></div>
      <div class="communication-mobile-field communication-mobile-field-full"><small>What Happened?</small><div class="communication-mobile-message">${escape(record.whatHappened || "-")}</div></div>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" class="ghost" data-guests-blacklist-action="edit" data-id="${escape(record.id)}">Edit</button><button type="button" class="danger" data-guests-blacklist-action="delete" data-id="${escape(record.id)}">Delete</button></div></div>`;
  return card;
}

function buildGuestsBlacklistEditableCard(record) {
  const draft = state.guestsBlacklistEditDraft || record;
  const card = document.createElement("article");
  card.className = "guests-mobile-card";
  card.innerHTML = `<div class="communication-mobile-grid">
      <label class="communication-mobile-field communication-mobile-field-full"><small>Name</small><input data-field="name" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.name)}" /></label>
      <label class="communication-mobile-field"><small>Nationality</small><input data-field="nationality" data-scope="edit" data-id="${escape(record.id)}" list="guests-country-list" value="${escape(draft.nationality)}" /></label>
      <label class="communication-mobile-field"><small>Birth Date</small><input data-field="birthDate" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.birthDate)}" /></label>
      <label class="communication-mobile-field"><small>Doc. Number</small><input data-field="docNumber" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.docNumber)}" /></label>
      <label class="communication-mobile-field"><small>Occurrence Date</small><input data-field="occurrenceDate" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.occurrenceDate)}" /></label>
      <label class="communication-mobile-field communication-mobile-field-full"><small>Who Reported</small><input data-field="whoReported" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.whoReported)}" /></label>
      <label class="communication-mobile-field communication-mobile-field-full"><small>What Happened?</small><textarea data-field="whatHappened" data-scope="edit" data-id="${escape(record.id)}" rows="3">${escape(draft.whatHappened)}</textarea></label>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" data-guests-blacklist-action="save-edit" data-id="${escape(record.id)}">Save</button><button type="button" class="ghost" data-guests-blacklist-action="cancel-edit" data-id="${escape(record.id)}">Cancel</button></div></div>`;
  return card;
}

function guestDescriptionPaletteClient(colorKey) {
  return GUEST_DESCRIPTION_PALETTES[clean(colorKey).toLowerCase()] || GUEST_DESCRIPTION_PALETTES.blue;
}

function buildGuestDescriptionRow(record, { showRoom = true, rowSpan = 1 } = {}) {
  const palette = guestDescriptionPaletteClient(record.colorKey);
  const tr = document.createElement("tr");
  tr.className = "guest-description-row";
  tr.dataset.colorKey = record.colorKey;
  const roomCell = showRoom
    ? `<td class="guest-description-fixed-cell" rowspan="${Math.max(1, Number(rowSpan) || 1)}" style="background:${palette.solid}">${escape(record.room || "-")}</td>`
    : "";
  tr.innerHTML = `${roomCell}
    <td class="guest-description-fixed-cell" style="background:${palette.solid}">${escape(record.bed || "-")}</td>
    <td class="guest-description-edit-cell" style="background:${palette.soft}"><textarea class="guest-description-input" data-guest-description-id="${escape(record.id)}" data-original-description="${escape(record.guestDescription)}" rows="2">${escape(record.guestDescription)}</textarea></td>`;
  return tr;
}

function buildGuestDescriptionCard(record) {
  const palette = guestDescriptionPaletteClient(record.colorKey);
  const card = document.createElement("article");
  card.className = "guests-mobile-card guest-description-card";
  card.style.background = palette.soft;
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(record.room || "-")}</div>
        <div class="communication-mobile-meta">Bed: ${escape(record.bed || "-")}</div>
      </div>
    </div>
    <div class="communication-mobile-grid">
      <label class="communication-mobile-field communication-mobile-field-full"><small>Guest Description</small><textarea class="guest-description-input" data-guest-description-id="${escape(record.id)}" data-original-description="${escape(record.guestDescription)}" rows="4">${escape(record.guestDescription)}</textarea></label>
    </div>`;
  return card;
}

function renderGuestDescriptionMobileCards(rows) {
  if (!els.guestsDescriptionsMobileCards) return;
  els.guestsDescriptionsMobileCards.innerHTML = "";
  if (!rows.length) {
    els.guestsDescriptionsMobileCards.innerHTML = '<div class="services-mobile-empty">No guest description rows found.</div>';
    return;
  }
  rows.forEach((record) => {
    els.guestsDescriptionsMobileCards.appendChild(buildGuestDescriptionCard(record));
  });
}

async function saveGuestDescriptionRow(id, guestDescription) {
  const result = await api(`/api/guests-descriptions?id=${encodeURIComponent(id)}`, {
    method: "PUT",
    body: { guestDescription },
  });
  state.guestDescriptionRows = sortGuestDescriptionRowsClient((Array.isArray(result?.rows) ? result.rows : []).map(normalizeGuestDescriptionRowClient));
  renderGuests();
}

function onGuestDescriptionFocusIn(event) {
  const target = event.target;
  if (!(target instanceof HTMLTextAreaElement) || !target.matches("[data-guest-description-id]")) return;
  target.dataset.originalDescription = target.value;
}

function onGuestDescriptionInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLTextAreaElement) || !target.matches("[data-guest-description-id]")) return;
  const id = clean(target.dataset.guestDescriptionId);
  const row = state.guestDescriptionRows.find((item) => item.id === id);
  if (!row) return;
  row.guestDescription = target.value;
}

function onGuestDescriptionsFilterInput() {
  state.guestsDescriptionsFilters.room = clean(els.guestsDescriptionsFilterRoom?.value);
  state.guestsDescriptionsFilters.description = clean(els.guestsDescriptionsFilterDescription?.value);
  renderGuests();
}

function getFilteredGuestDescriptionRows() {
  const filters = state.guestsDescriptionsFilters || {};
  const room = clean(filters.room).toLowerCase();
  const description = clean(filters.description).toLowerCase();
  return sortGuestDescriptionRowsClient(state.guestDescriptionRows).filter((row) => {
    const roomMatch = !room || clean(row.room).toLowerCase().includes(room);
    const descriptionMatch = !description || clean(row.guestDescription).toLowerCase().includes(description);
    return roomMatch && descriptionMatch;
  });
}

function buildGuestDescriptionTableRows(rows) {
  const records = Array.isArray(rows) ? rows : [];
  const builtRows = [];
  let index = 0;
  while (index < records.length) {
    const current = records[index];
    const room = clean(current?.room);
    let rowSpan = 1;
    while (index + rowSpan < records.length && clean(records[index + rowSpan]?.room) === room) {
      rowSpan += 1;
    }
    builtRows.push(buildGuestDescriptionRow(current, { showRoom: true, rowSpan }));
    for (let offset = 1; offset < rowSpan; offset += 1) {
      builtRows.push(buildGuestDescriptionRow(records[index + offset], { showRoom: false }));
    }
    index += rowSpan;
  }
  return builtRows;
}

async function onGuestDescriptionFocusOut(event) {
  const target = event.target;
  if (!(target instanceof HTMLTextAreaElement) || !target.matches("[data-guest-description-id]")) return;
  const id = clean(target.dataset.guestDescriptionId);
  const nextValue = clean(target.value);
  const previousValue = clean(target.dataset.originalDescription);
  if (!id || nextValue === previousValue) return;
  try {
    setGuestsDescriptionsStatus("Saving guest description...");
    await saveGuestDescriptionRow(id, nextValue);
    setGuestsDescriptionsStatus("Guest description saved.");
  } catch (e) {
    const row = state.guestDescriptionRows.find((item) => item.id === id);
    const fallback = previousValue;
    if (row) row.guestDescription = fallback;
    target.value = fallback;
    target.dataset.originalDescription = fallback;
    setGuestsDescriptionsStatus(`Save failed: ${e.message}`);
    showToast(`Guest description save failed: ${e.message}`, "error");
  }
}

function renderGuestsMobileCards(rows) {
  if (!els.guestsMobileCards) return;
  els.guestsMobileCards.innerHTML = "";
  els.guestsMobileCards.appendChild(buildGuestsInlineCard());
  if (!rows.length) {
    els.guestsMobileCards.innerHTML += '<div class="services-mobile-empty">No guest records found.</div>';
    return;
  }
  rows.forEach((record) => {
    els.guestsMobileCards.appendChild(state.guestsEditingId === record.id ? buildGuestsEditableCard(record) : buildGuestsReadOnlyCard(record));
  });
}

function renderGuestsBlacklistMobileCards(rows) {
  if (!els.guestsBlacklistMobileCards) return;
  els.guestsBlacklistMobileCards.innerHTML = "";
  els.guestsBlacklistMobileCards.appendChild(buildGuestsBlacklistInlineCard());
  if (!rows.length) {
    els.guestsBlacklistMobileCards.innerHTML += '<div class="services-mobile-empty">No blacklist records found.</div>';
    return;
  }
  rows.forEach((record) => {
    els.guestsBlacklistMobileCards.appendChild(state.guestsBlacklistEditingId === record.id ? buildGuestsBlacklistEditableCard(record) : buildGuestsBlacklistReadOnlyCard(record));
  });
}

function renderGuestsScreenTabs() {
  if (els.guestsTabList) {
    els.guestsTabList.classList.toggle("active-tab", state.guestsScreen === "list");
    els.guestsTabList.classList.toggle("ghost", state.guestsScreen !== "list");
  }
  if (els.guestsTabDescriptions) {
    els.guestsTabDescriptions.classList.toggle("active-tab", state.guestsScreen === "descriptions");
    els.guestsTabDescriptions.classList.toggle("ghost", state.guestsScreen !== "descriptions");
  }
  if (els.guestsTabBlacklist) {
    els.guestsTabBlacklist.classList.toggle("active-tab", state.guestsScreen === "blacklist");
    els.guestsTabBlacklist.classList.toggle("ghost", state.guestsScreen !== "blacklist");
  }
  if (els.guestsPanelList) els.guestsPanelList.hidden = state.guestsScreen !== "list";
  if (els.guestsPanelDescriptions) els.guestsPanelDescriptions.hidden = state.guestsScreen !== "descriptions";
  if (els.guestsPanelBlacklist) els.guestsPanelBlacklist.hidden = state.guestsScreen !== "blacklist";
}

function shouldShowGuestsAlertClient(todayIso = lisbonTodayIsoClient()) {
  if (!canApp("guests")) return false;
  const sendTime = clean(state.guestsSettings?.sendTime) || DEFAULT_GUESTS_SETTINGS.sendTime;
  const currentTime = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/Lisbon",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(new Date());
  if (currentTime < sendTime) return false;
  return state.guestsRows.some((row) => clean(row.sentStatus).toLowerCase() !== "sent" && clean(row.checkIn) && clean(row.checkIn) < todayIso);
}

function getGuestsAlertReasonText(todayIso = lisbonTodayIsoClient()) {
  const affected = state.guestsRows.filter((row) => clean(row.sentStatus).toLowerCase() !== "sent" && clean(row.checkIn) && clean(row.checkIn) < todayIso);
  if (!affected.length) return "";
  return `Pending send: ${affected.length} guest record${affected.length === 1 ? "" : "s"} with check-in before today.`;
}

function renderGuests() {
  renderGuestsScreenTabs();
  renderGuestsCountryOptions();
  if (!canApp("guests")) {
    if (els.guestsCount) els.guestsCount.textContent = "0 records";
    if (els.guestsAlertSummary) els.guestsAlertSummary.textContent = "";
    if (els.guestsListAlertReason) els.guestsListAlertReason.textContent = "";
    if (els.guestsRows) els.guestsRows.innerHTML = '<tr><td colspan="12" class="empty">Your profile has no access to Guests.</td></tr>';
    if (els.guestsMobileCards) els.guestsMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Guests.</div>';
    if (els.guestsDescriptionsCount) els.guestsDescriptionsCount.textContent = "0 rows";
    if (els.guestsDescriptionsRows) els.guestsDescriptionsRows.innerHTML = '<tr><td colspan="3" class="empty">Your profile has no access to Guests.</td></tr>';
    if (els.guestsDescriptionsMobileCards) els.guestsDescriptionsMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Guests.</div>';
    if (els.guestsBlacklistRows) els.guestsBlacklistRows.innerHTML = '<tr><td colspan="8" class="empty">Your profile has no access to Guests.</td></tr>';
    if (els.guestsBlacklistMobileCards) els.guestsBlacklistMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Guests.</div>';
    return;
  }
  if (els.guestsShowActive) els.guestsShowActive.checked = !!state.guestsFilters.showActive;
  if (els.guestsFilterHa) els.guestsFilterHa.value = state.guestsFilters.ha;
  if (els.guestsFilterSearch) els.guestsFilterSearch.value = state.guestsFilters.search;
  if (els.guestsFilterNationality) els.guestsFilterNationality.value = state.guestsFilters.nationality;
  if (els.guestsFilterCheckinFrom) els.guestsFilterCheckinFrom.value = state.guestsFilters.checkInFrom;
  if (els.guestsFilterCheckinTo) els.guestsFilterCheckinTo.value = state.guestsFilters.checkInTo;
  if (els.guestsFilterCheckoutFrom) els.guestsFilterCheckoutFrom.value = state.guestsFilters.checkOutFrom;
  if (els.guestsFilterCheckoutTo) els.guestsFilterCheckoutTo.value = state.guestsFilters.checkOutTo;
  if (els.guestsDescriptionsFilterRoom) els.guestsDescriptionsFilterRoom.value = state.guestsDescriptionsFilters.room;
  if (els.guestsDescriptionsFilterDescription) els.guestsDescriptionsFilterDescription.value = state.guestsDescriptionsFilters.description;
  if (els.guestsBlacklistFilterSearch) els.guestsBlacklistFilterSearch.value = state.guestsBlacklistFilters.search;
  if (els.guestsBlacklistFilterReported) els.guestsBlacklistFilterReported.value = state.guestsBlacklistFilters.whoReported;
  if (els.guestsBlacklistFilterNationality) els.guestsBlacklistFilterNationality.value = state.guestsBlacklistFilters.nationality;
  const rows = getFilteredGuestsRows();
  const descriptionRows = getFilteredGuestDescriptionRows();
  const blacklistRows = getFilteredGuestsBlacklistRows();
  if (els.guestsListAlertReason) els.guestsListAlertReason.textContent = shouldShowGuestsAlertClient() ? getGuestsAlertReasonText() : "";
  if (els.guestsAlertSummary) els.guestsAlertSummary.textContent = getGuestsTopAlertsSummaryText();
  const sendableCount = state.guestsRows.filter((row) => clean(row.sentStatus).toLowerCase() !== "sent" && clean(row.checkIn) && clean(row.checkIn) <= lisbonTodayIsoClient()).length;
  if (els.guestsSendPending) els.guestsSendPending.disabled = sendableCount === 0;
  if (els.guestsCount) els.guestsCount.textContent = `${rows.length} record${rows.length === 1 ? "" : "s"}`;
  if (els.guestsDescriptionsCount) els.guestsDescriptionsCount.textContent = `${descriptionRows.length} row${descriptionRows.length === 1 ? "" : "s"}`;
  if (els.guestsBlacklistCount) els.guestsBlacklistCount.textContent = `${blacklistRows.length} record${blacklistRows.length === 1 ? "" : "s"}`;
  if (state.guestsScreen === "descriptions") {
    if (els.guestsDescriptionsRows) els.guestsDescriptionsRows.innerHTML = "";
    renderGuestDescriptionMobileCards(descriptionRows);
    if (!els.guestsDescriptionsRows) return;
    if (!descriptionRows.length) {
      const tr = document.createElement("tr");
      tr.innerHTML = '<td colspan="3" class="empty">No guest description rows found.</td>';
      els.guestsDescriptionsRows.appendChild(tr);
      return;
    }
    buildGuestDescriptionTableRows(descriptionRows).forEach((row) => {
      els.guestsDescriptionsRows.appendChild(row);
    });
    return;
  }
  if (state.guestsScreen === "blacklist") {
    if (els.guestsBlacklistRows) els.guestsBlacklistRows.innerHTML = "";
    renderGuestsBlacklistMobileCards(blacklistRows);
    if (!els.guestsBlacklistRows) return;
    els.guestsBlacklistRows.appendChild(buildGuestsBlacklistInlineRow());
    if (!blacklistRows.length) {
      const tr = document.createElement("tr");
      tr.innerHTML = '<td colspan="8" class="empty">No blacklist records found.</td>';
      els.guestsBlacklistRows.appendChild(tr);
      return;
    }
    blacklistRows.forEach((record) => {
      els.guestsBlacklistRows.appendChild(state.guestsBlacklistEditingId === record.id ? buildGuestsBlacklistEditableRow(record) : buildGuestsBlacklistReadOnlyRow(record));
    });
    return;
  }
  if (els.guestsRows) els.guestsRows.innerHTML = "";
  renderGuestsMobileCards(rows);
  if (!els.guestsRows) return;
  els.guestsRows.appendChild(buildGuestsInlineRow());
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td colspan="12" class="empty">No guest records found.</td>';
    els.guestsRows.appendChild(tr);
    return;
  }
  rows.forEach((record) => {
    els.guestsRows.appendChild(state.guestsEditingId === record.id ? buildGuestsEditableRow(record) : buildGuestsReadOnlyRow(record));
  });
}

function hasGuestsDraft() {
  const draft = state.guestsEditingId ? state.guestsEditDraft : state.guestsDraft;
  return !!(
    clean(draft?.name) ||
    clean(draft?.nationality) ||
    clean(draft?.birthDate) ||
    clean(draft?.docNumber) ||
    clean(draft?.checkIn) ||
    clean(draft?.checkOut)
  );
}

function hasGuestsBlacklistDraft() {
  const draft = state.guestsBlacklistEditingId ? state.guestsBlacklistEditDraft : state.guestsBlacklistDraft;
  return !!(
    clean(draft?.name) ||
    clean(draft?.docNumber) ||
    clean(draft?.occurrenceDate) ||
    clean(draft?.whatHappened)
  );
}

async function loadHoursSettings({ silent = false } = {}) {
  try {
    const result = await api("/api/hours-register-settings");
    state.hoursSettings = normalizeHoursSettingsClient(result?.settings);
    state.hoursSettingsLoaded = true;
    renderHoursSettings();
    if (!silent) setHoursSettingsStatus("Hours configuration loaded.");
  } catch (e) {
    state.hoursSettings = clone(DEFAULT_HOURS_SETTINGS);
    renderHoursSettings();
    if (!silent) setHoursSettingsStatus(`Using default hours configuration (${e.message}).`);
  }
}

async function loadHoursData({ silent = false } = {}) {
  try {
    const result = await api("/api/hours-register");
    state.hoursSettings = normalizeHoursSettingsClient(result?.settings || state.hoursSettings);
    state.hoursRecords = sortHoursRecordsClient((Array.isArray(result?.rows) ? result.rows : []).map((row) => normalizeHoursRecordClient(row, state.hoursSettings)));
    state.hoursLoaded = true;
    if (!state.hoursEditingId) state.hoursDraft = emptyHoursDraft();
    renderHours();
    renderHoursSettings();
    if (!silent) setHoursStatus("Hours records loaded.");
  } catch (e) {
    state.hoursRecords = [];
    state.hoursSettings = clone(DEFAULT_HOURS_SETTINGS);
    state.hoursDraft = emptyHoursDraft();
    renderHours();
    renderHoursSettings();
    if (!silent) setHoursStatus(`Using default hours data (${e.message}).`);
  }
}

function onHoursSettingsInput() {
  state.hoursSettings = normalizeHoursSettingsClient({
    people: String(els.hoursSettingsPersons?.value || "").split(/\r?\n/),
  });
  if (!state.hoursEditingId && !clean(state.hoursDraft?.person) && state.hoursSettings.people.length === 1) {
    state.hoursDraft = { ...(state.hoursDraft || emptyHoursDraft()), person: state.hoursSettings.people[0] };
  }
}

async function saveHoursSettings() {
  onHoursSettingsInput();
  if (!state.hoursSettings.people.length) {
    setHoursSettingsStatus("At least one person is required.");
    return;
  }
  try {
    const result = await api("/api/hours-register-settings", { method: "PUT", body: { settings: state.hoursSettings } });
    state.hoursSettings = normalizeHoursSettingsClient(result?.settings);
    state.hoursSettingsLoaded = true;
    state.hoursRecords = state.hoursRecords.map((row) => normalizeHoursRecordClient(row, state.hoursSettings));
    if (!state.hoursEditingId) state.hoursDraft = emptyHoursDraft();
    renderHoursSettings();
    renderHours();
    setHoursSettingsStatus("Hours configuration saved.");
    showToast("Hours configuration saved.", "success");
  } catch (e) {
    setHoursSettingsStatus(`Save failed: ${e.message}`);
    showToast(`Hours configuration save failed: ${e.message}`, "error");
  }
}

function renderHoursSettings() {
  const settings = state.hoursSettings || clone(DEFAULT_HOURS_SETTINGS);
  if (els.hoursSettingsPersons) els.hoursSettingsPersons.value = (settings.people || []).join("\n");
}

function setHoursScreen(screen) {
  state.hoursScreen = screen === "resume" ? "resume" : "list";
  renderHours();
}

function onHoursFilterInput(event) {
  const target = event?.target;
  if (target === els.hoursFilterPerson || target === els.hoursResumeFilterPerson) state.hoursFilters.person = clean(target.value);
  if (target === els.hoursFilterDateFrom || target === els.hoursResumeFilterDateFrom) state.hoursFilters.dateFrom = clean(target.value);
  if (target === els.hoursFilterDateTo || target === els.hoursResumeFilterDateTo) state.hoursFilters.dateTo = clean(target.value);
  renderHours();
}

function getFilteredHoursRecords() {
  const filters = state.hoursFilters || {};
  const person = clean(filters.person).toLowerCase();
  const dateFrom = clean(filters.dateFrom);
  const dateTo = clean(filters.dateTo);
  return sortHoursRecordsClient(state.hoursRecords)
    .filter((row) => !person || clean(row.person).toLowerCase() === person)
    .filter((row) => !dateFrom || clean(row.date) >= dateFrom)
    .filter((row) => !dateTo || clean(row.date) <= dateTo);
}

function renderHoursPersonOptions() {
  const people = state.hoursSettings?.people || DEFAULT_HOURS_SETTINGS.people;
  const allOptions = [`<option value="">All</option>`, ...people.map((person) => option(person, ""))].join("");
  if (els.hoursFilterPerson) {
    els.hoursFilterPerson.innerHTML = allOptions;
    els.hoursFilterPerson.value = state.hoursFilters.person;
  }
  if (els.hoursResumeFilterPerson) {
    els.hoursResumeFilterPerson.innerHTML = allOptions;
    els.hoursResumeFilterPerson.value = state.hoursFilters.person;
  }
}

function validateHoursDraftClient(draft) {
  const person = clean(draft?.person);
  if (!person) return "Person is required.";
  if (!(state.hoursSettings?.people || []).some((item) => clean(item).toLowerCase() === person.toLowerCase())) return "Person must exist in the configured list.";
  if (!clean(draft?.date)) return "Date is required.";
  if (!clean(draft?.start)) return "Start time is required.";
  if (clean(draft?.finish)) {
    const minutes = hoursDurationMinutes(draft.start, draft.finish);
    if (minutes == null || minutes <= 0) return "Finish time must be after start time.";
  }
  return "";
}

function buildHoursPayload(draft) {
  return {
    person: clean(draft.person),
    date: normalizeDateInput(draft.date),
    start: normalizeHoursTimeValue(draft.start),
    finish: normalizeHoursTimeValue(draft.finish),
  };
}

async function saveHoursRecord(mode = "new", id = "") {
  const isEdit = mode === "edit";
  const draft = isEdit ? state.hoursEditDraft : state.hoursDraft;
  const pending = !isEdit ? getPendingHoursRecord() : null;
  if (pending) {
    const message = `Please fill finish time for ${pending.person} on ${pending.date} before adding a new record.`;
    setHoursStatus(message);
    showToast(message, "error");
    return;
  }
  const validationError = validateHoursDraftClient(draft);
  if (validationError) {
    setHoursStatus(validationError);
    showToast(validationError, "error");
    return;
  }
  try {
    const result = await api(isEdit ? `/api/hours-register?id=${encodeURIComponent(id)}` : "/api/hours-register", {
      method: isEdit ? "PUT" : "POST",
      body: buildHoursPayload(draft),
    });
    state.hoursSettings = normalizeHoursSettingsClient(result?.settings || state.hoursSettings);
    state.hoursRecords = sortHoursRecordsClient((Array.isArray(result?.rows) ? result.rows : []).map((row) => normalizeHoursRecordClient(row, state.hoursSettings)));
    state.hoursLoaded = true;
    state.hoursEditingId = null;
    state.hoursEditDraft = null;
    state.hoursDraft = emptyHoursDraft();
    renderHours();
    renderHoursSettings();
    setHoursStatus(isEdit ? "Hours record saved." : "Hours record added.");
    showToast(isEdit ? "Hours record saved." : "Hours record added.", "success");
  } catch (e) {
    setHoursStatus(`Save failed: ${e.message}`);
    showToast(`Hours record save failed: ${e.message}`, "error");
  }
}

function onHoursDraftInput(event) {
  const field = clean(event.target.dataset.field);
  const scope = clean(event.target.dataset.scope || "new");
  if (!field) return;
  const targetDraft = scope === "edit" ? state.hoursEditDraft : state.hoursDraft;
  if (!targetDraft) return;
  targetDraft[field] = event.target.value;
  if ((field === "person" || field === "date" || field === "start" || field === "finish") && els.hoursStatus) {
    setHoursStatus("");
  }
  renderHours();
}

async function onHoursAction(event) {
  const button = event.target.closest("button[data-hours-action]");
  if (!button) return;
  const action = clean(button.dataset.hoursAction);
  const id = clean(button.dataset.id);
  if (action === "save-inline") {
    await saveHoursRecord("new");
    return;
  }
  if (action === "edit" && id) {
    await loadHoursData({ silent: true });
    const record = state.hoursRecords.find((item) => clean(item.id) === id);
    if (!record) return;
    state.hoursEditingId = id;
    state.hoursEditDraft = { ...record };
    renderHours();
    return;
  }
  if (action === "save-edit" && id) {
    await saveHoursRecord("edit", id);
    return;
  }
  if (action === "cancel-edit") {
    state.hoursEditingId = null;
    state.hoursEditDraft = null;
    renderHours();
  }
}

function buildHoursInlineRow() {
  const draft = state.hoursDraft || emptyHoursDraft();
  const pending = getPendingHoursRecord();
  const tr = document.createElement("tr");
  tr.className = "inline-editor sticky-new-row";
  tr.innerHTML = `<td><select data-field="person" data-scope="new">${(state.hoursSettings?.people || DEFAULT_HOURS_SETTINGS.people).map((person) => option(person, draft.person)).join("")}</select></td>
    <td><input data-field="date" data-scope="new" type="date" value="${escape(draft.date)}" /></td>
    <td><input data-field="start" data-scope="new" type="time" value="${escape(draft.start)}" /></td>
    <td><input data-field="finish" data-scope="new" type="time" value="${escape(draft.finish)}" /></td>
    <td>${escape(formatHoursDuration(draft.start, draft.finish))}</td>
    <td class="row-actions"><button type="button" data-hours-action="save-inline" ${pending ? "disabled" : ""}>Add</button>${pending ? '<div class="warning-text hours-pending-inline">please fill finish time</div>' : ""}</td>`;
  return tr;
}

function buildHoursReadOnlyRow(record) {
  const tr = document.createElement("tr");
  tr.innerHTML = `<td>${escape(record.person)}</td>
    <td>${escape(record.date)}</td>
    <td>${escape(record.start)}</td>
    <td>${escape(record.finish || "-")}${hoursRecordNeedsFinish(record) ? '<br><span class="warning-text">please fill finish time</span>' : ""}</td>
    <td>${escape(formatHoursDuration(record.start, record.finish))}</td>
    <td class="row-actions"><button type="button" class="ghost" data-hours-action="edit" data-id="${escape(record.id)}">Edit</button></td>`;
  return tr;
}

function buildHoursEditableRow(record) {
  const draft = state.hoursEditDraft || record;
  const tr = document.createElement("tr");
  tr.className = "inline-editor";
  tr.innerHTML = `<td><select data-field="person" data-scope="edit" data-id="${escape(record.id)}">${(state.hoursSettings?.people || DEFAULT_HOURS_SETTINGS.people).map((person) => option(person, draft.person)).join("")}</select></td>
    <td><input data-field="date" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.date)}" /></td>
    <td><input data-field="start" data-scope="edit" data-id="${escape(record.id)}" type="time" value="${escape(draft.start)}" /></td>
    <td><input data-field="finish" data-scope="edit" data-id="${escape(record.id)}" type="time" value="${escape(draft.finish)}" /></td>
    <td>${escape(formatHoursDuration(draft.start, draft.finish))}</td>
    <td class="row-actions"><button type="button" data-hours-action="save-edit" data-id="${escape(record.id)}">Save</button>
    <button type="button" class="ghost" data-hours-action="cancel-edit" data-id="${escape(record.id)}">Cancel</button></td>`;
  return tr;
}

function buildHoursInlineCard() {
  const draft = state.hoursDraft || emptyHoursDraft();
  const pending = getPendingHoursRecord();
  const card = document.createElement("article");
  card.className = "hours-mobile-card";
  card.innerHTML = `<div class="communication-mobile-grid">
      <label class="communication-mobile-field">
        <small>Person</small>
        <select data-field="person" data-scope="new">${(state.hoursSettings?.people || DEFAULT_HOURS_SETTINGS.people).map((person) => option(person, draft.person)).join("")}</select>
      </label>
      <label class="communication-mobile-field">
        <small>Date</small>
        <input data-field="date" data-scope="new" type="date" value="${escape(draft.date)}" />
      </label>
      <label class="communication-mobile-field">
        <small>Start</small>
        <input data-field="start" data-scope="new" type="time" value="${escape(draft.start)}" />
      </label>
      <label class="communication-mobile-field">
        <small>Finish</small>
        <input data-field="finish" data-scope="new" type="time" value="${escape(draft.finish)}" />
      </label>
      <div class="communication-mobile-field communication-mobile-field-full">
        <small>Hours</small>
        <div class="communication-mobile-message">${escape(formatHoursDuration(draft.start, draft.finish))}</div>
      </div>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" data-hours-action="save-inline" ${pending ? "disabled" : ""}>Add</button></div>${pending ? '<span class="warning-text">please fill finish time</span>' : ""}</div>`;
  return card;
}

function buildHoursReadOnlyCard(record) {
  const card = document.createElement("article");
  card.className = "hours-mobile-card";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(record.person)}</div>
        <div class="communication-mobile-meta">${escape(record.date)}</div>
      </div>
      <div class="group-mobile-total">
        <strong>${escape(formatHoursDuration(record.start, record.finish))}</strong>
        <small>Hours</small>
      </div>
    </div>
    <div class="communication-mobile-grid">
      <div class="communication-mobile-field"><small>Start</small><div class="communication-mobile-message">${escape(record.start)}</div></div>
      <div class="communication-mobile-field"><small>Finish</small><div class="communication-mobile-message">${escape(record.finish || "-")}${hoursRecordNeedsFinish(record) ? '<br><span class="warning-text">please fill finish time</span>' : ""}</div></div>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions"><button type="button" class="ghost" data-hours-action="edit" data-id="${escape(record.id)}">Edit</button></div></div>`;
  return card;
}

function buildHoursEditableCard(record) {
  const draft = state.hoursEditDraft || record;
  const card = document.createElement("article");
  card.className = "hours-mobile-card";
  card.innerHTML = `<div class="communication-mobile-grid">
      <label class="communication-mobile-field">
        <small>Person</small>
        <select data-field="person" data-scope="edit" data-id="${escape(record.id)}">${(state.hoursSettings?.people || DEFAULT_HOURS_SETTINGS.people).map((person) => option(person, draft.person)).join("")}</select>
      </label>
      <label class="communication-mobile-field">
        <small>Date</small>
        <input data-field="date" data-scope="edit" data-id="${escape(record.id)}" type="date" value="${escape(draft.date)}" />
      </label>
      <label class="communication-mobile-field">
        <small>Start</small>
        <input data-field="start" data-scope="edit" data-id="${escape(record.id)}" type="time" value="${escape(draft.start)}" />
      </label>
      <label class="communication-mobile-field">
        <small>Finish</small>
        <input data-field="finish" data-scope="edit" data-id="${escape(record.id)}" type="time" value="${escape(draft.finish)}" />
      </label>
      <div class="communication-mobile-field communication-mobile-field-full">
        <small>Hours</small>
        <div class="communication-mobile-message">${escape(formatHoursDuration(draft.start, draft.finish))}</div>
      </div>
    </div>
    <div class="communication-mobile-footer"><div class="row-actions">
      <button type="button" data-hours-action="save-edit" data-id="${escape(record.id)}">Save</button>
      <button type="button" class="ghost" data-hours-action="cancel-edit" data-id="${escape(record.id)}">Cancel</button>
    </div></div>`;
  return card;
}

function renderHoursMobileCards(rows) {
  if (!els.hoursMobileCards) return;
  const list = els.hoursMobileCards;
  list.innerHTML = "";
  list.appendChild(buildHoursInlineCard());
  if (!rows.length) {
    list.innerHTML += '<div class="services-mobile-empty">No hours records found.</div>';
    return;
  }
  rows.forEach((record) => {
    list.appendChild(state.hoursEditingId === record.id ? buildHoursEditableCard(record) : buildHoursReadOnlyCard(record));
  });
}

function getHoursResumeRows() {
  const buckets = new Map();
  getFilteredHoursRecords().filter((row) => !hoursRecordNeedsFinish(row)).forEach((row) => {
    const monthKey = clean(row.date).slice(0, 7);
    if (!monthKey) return;
    const key = `${monthKey}::${clean(row.person).toLowerCase()}`;
    const durationMinutes = Math.max(0, hoursDurationMinutes(row.start, row.finish) || 0);
    const current = buckets.get(key) || { monthKey, person: row.person, hours: 0, totalMinutes: 0, records: 0 };
    current.hours += durationMinutes / 60;
    current.totalMinutes += durationMinutes;
    current.records += 1;
    buckets.set(key, current);
  });
  return [...buckets.values()]
    .map((row) => ({
      ...row,
      hours: Number(row.hours.toFixed(2)),
      totalMinutes: Math.round(row.totalMinutes),
      averageHours: row.records ? Number((row.hours / row.records).toFixed(2)) : 0,
      averageMinutes: row.records ? Math.round(row.totalMinutes / row.records) : 0,
    }))
    .sort((a, b) => {
      const monthCompare = clean(b.monthKey).localeCompare(clean(a.monthKey));
      if (monthCompare !== 0) return monthCompare;
      return clean(a.person).localeCompare(clean(b.person));
    });
}

function buildHoursExportRows() {
  return getFilteredHoursRecords().map((row) => ({
    person: clean(row.person),
    date: clean(row.date),
    start: clean(row.start),
    finish: clean(row.finish),
    hours: formatHoursDuration(row.start, row.finish),
  }));
}

function exportHoursToExcel() {
  const rows = buildHoursExportRows();
  if (!rows.length) {
    showToast("No hours records to export.", "error");
    return;
  }
  const headers = ["Person", "Date", "Start", "Finish", "Hours"];
  const bodyRows = rows.map((row) => [row.person, row.date, row.start, row.finish, row.hours]);
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><table border="1"><thead><tr>${headers.map((cell) => `<th>${escape(cell)}</th>`).join("")}</tr></thead><tbody>${bodyRows.map((cells) => `<tr>${cells.map((cell) => `<td>${escape(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table></body></html>`;
  downloadBlob(`hours_register_${formatDate(new Date())}.xls`, html, "application/vnd.ms-excel;charset=utf-8;");
  showToast(`Exported ${rows.length} hours record${rows.length === 1 ? "" : "s"} to Excel.`, "success");
}

function formatHoursMonthLabel(monthKey) {
  if (!/^\d{4}-\d{2}$/.test(clean(monthKey))) return monthKey || "-";
  const dt = new Date(`${monthKey}-01T00:00:00`);
  if (Number.isNaN(dt.getTime())) return monthKey;
  return new Intl.DateTimeFormat("en-GB", { month: "short", year: "numeric", timeZone: "Europe/Lisbon" }).format(dt);
}

function renderHoursResume() {
  const rows = getHoursResumeRows();
  if (els.hoursResumeCount) els.hoursResumeCount.textContent = `${rows.length} row${rows.length === 1 ? "" : "s"}`;
  if (els.hoursResumeFilterPerson) els.hoursResumeFilterPerson.value = state.hoursFilters.person;
  if (els.hoursResumeFilterDateFrom) els.hoursResumeFilterDateFrom.value = state.hoursFilters.dateFrom;
  if (els.hoursResumeFilterDateTo) els.hoursResumeFilterDateTo.value = state.hoursFilters.dateTo;
  if (!els.hoursResumeBody) return;
  els.hoursResumeBody.innerHTML = "";
  if (!rows.length) {
    els.hoursResumeBody.innerHTML = '<tr><td colspan="5" class="empty">No hours summary found.</td></tr>';
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(formatHoursMonthLabel(row.monthKey))}</td><td>${escape(row.person)}</td><td>${escape(String(row.records))}</td><td>${escape(formatHoursMinutes(row.averageMinutes))}</td><td>${escape(formatHoursMinutes(row.totalMinutes))}</td>`;
    els.hoursResumeBody.appendChild(tr);
  });
}

function renderHoursScreenTabs() {
  if (els.hoursTabList) {
    els.hoursTabList.classList.toggle("active-tab", state.hoursScreen === "list");
    els.hoursTabList.classList.toggle("ghost", state.hoursScreen !== "list");
  }
  if (els.hoursTabResume) {
    els.hoursTabResume.classList.toggle("active-tab", state.hoursScreen === "resume");
    els.hoursTabResume.classList.toggle("ghost", state.hoursScreen !== "resume");
  }
  if (els.hoursPanelList) els.hoursPanelList.hidden = state.hoursScreen !== "list";
  if (els.hoursPanelResume) els.hoursPanelResume.hidden = state.hoursScreen !== "resume";
}

function renderHours() {
  if (!canApp("hours")) {
    if (els.hoursCount) els.hoursCount.textContent = "0 records";
    if (els.hoursRows) els.hoursRows.innerHTML = '<tr><td colspan="6" class="empty">Your profile has no access to Hours Register.</td></tr>';
    if (els.hoursMobileCards) els.hoursMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Hours Register.</div>';
    return;
  }
  renderHoursScreenTabs();
  renderHoursPersonOptions();
  if (state.hoursScreen === "resume") {
    renderHoursResume();
    return;
  }
  const rows = getFilteredHoursRecords();
  if (els.hoursCount) els.hoursCount.textContent = `${rows.length} record${rows.length === 1 ? "" : "s"}`;
  if (els.hoursFilterDateFrom) els.hoursFilterDateFrom.value = state.hoursFilters.dateFrom;
  if (els.hoursFilterDateTo) els.hoursFilterDateTo.value = state.hoursFilters.dateTo;
  if (els.hoursRows) els.hoursRows.innerHTML = "";
  renderHoursMobileCards(rows);
  if (!els.hoursRows) return;
  els.hoursRows.appendChild(buildHoursInlineRow());
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td colspan="6" class="empty">No hours records found.</td>';
    els.hoursRows.appendChild(tr);
    return;
  }
  rows.forEach((record) => {
    els.hoursRows.appendChild(state.hoursEditingId === record.id ? buildHoursEditableRow(record) : buildHoursReadOnlyRow(record));
  });
}

function hasHoursDraft() {
  const draft = state.hoursEditingId ? state.hoursEditDraft : state.hoursDraft;
  return !!(clean(draft?.person) || clean(draft?.date) || clean(draft?.start) || clean(draft?.finish));
}

function normalizeLaundrySettingsClient(input = {}) {
  const source = input && typeof input === "object" ? input : {};
  const itemTypes = (Array.isArray(source.itemTypes) ? source.itemTypes : [])
    .map((item, index) => ({
      id: clean(item?.id) || `laundry-item-${index + 1}`,
      name: clean(item?.name) || `item ${index + 1}`,
      weightKg: Math.max(0, Number(normalizeNumber(item?.weightKg) || 0)),
    }))
    .filter((item, index, items) => item.name && items.findIndex((candidate) => candidate.id === item.id) === index);
  return {
    pricePerKg: Math.max(0, Number(normalizeNumber(source.pricePerKg) || 0)),
    emailRecipients: String(Array.isArray(source.emailRecipients) ? source.emailRecipients.join("\n") : source.emailRecipients || "")
      .split(/[\n,;]/)
      .map((item) => clean(item).toLowerCase())
      .filter(Boolean)
      .filter((item, index, items) => items.indexOf(item) === index),
    emailEnabled: !!(source.emailEnabled ?? source.email_enabled),
    emailTime: normalizeTimeInput(source.emailTime ?? source.email_time),
    managementEmailRecipients: String(Array.isArray(source.managementEmailRecipients) ? source.managementEmailRecipients.join("\n") : source.managementEmailRecipients || source.management_email_recipients || "")
      .split(/[\n,;]/)
      .map((item) => clean(item).toLowerCase())
      .filter(Boolean)
      .filter((item, index, items) => items.indexOf(item) === index),
    managementEmailEnabled: !!(source.managementEmailEnabled ?? source.management_email_enabled),
    managementEmailTime: normalizeTimeInput(source.managementEmailTime ?? source.management_email_time),
    itemTypes: itemTypes.length ? itemTypes : clone(DEFAULT_LAUNDRY_SETTINGS.itemTypes),
  };
}

function normalizeLaundryPropertyClient(value) {
  const raw = clean(value).toLowerCase();
  if (!raw) return "Hostel";
  if (raw.includes("cruz") || raw.includes("apart")) return "Cruz";
  return "Hostel";
}

function sanitizeLaundryCountsClient(value, itemTypes = state.laundrySettings?.itemTypes || DEFAULT_LAUNDRY_SETTINGS.itemTypes) {
  const source = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  return itemTypes.reduce((acc, item) => {
    const raw = source[item.id];
    const normalized = normalizeNumber(raw);
    acc[item.id] = normalized == null ? null : Math.max(0, Math.round(Number(normalized || 0)));
    return acc;
  }, {});
}

function normalizeLaundryRecordClient(input = {}, settings = state.laundrySettings) {
  const safeSettings = normalizeLaundrySettingsClient(settings);
  const sentDate = clean(input.date);
  const rawReceivedWeight = normalizeNumber(input.receivedWeightKg);
  return {
    id: clean(input.id),
    property: normalizeLaundryPropertyClient(input.property),
    date: sentDate,
    receivedDate: clean(input.receivedDate) || laundryReceiveDate(sentDate),
    sentItems: sanitizeLaundryCountsClient(input.sentItems, safeSettings.itemTypes),
    receivedItems: sanitizeLaundryCountsClient(input.receivedItems, safeSettings.itemTypes),
    receivedWeightKg: rawReceivedWeight == null ? "" : Math.max(0, Number(rawReceivedWeight || 0)),
    notes: clean(input.notes),
    createdAt: clean(input.createdAt),
    updatedAt: clean(input.updatedAt),
  };
}

function emptyLaundryDraft() {
  const defaultProperty = "Hostel";
  const nextDate = getRequiredNextLaundryDate(defaultProperty) || formatDate(new Date());
  return normalizeLaundryRecordClient({
    id: "",
    property: defaultProperty,
    date: nextDate,
    sentItems: {},
    receivedItems: {},
    receivedWeightKg: "",
    notes: "",
    createdAt: "",
    updatedAt: "",
  });
}

function laundryItemTypes() {
  return Array.isArray(state.laundrySettings?.itemTypes) && state.laundrySettings.itemTypes.length
    ? state.laundrySettings.itemTypes
    : clone(DEFAULT_LAUNDRY_SETTINGS.itemTypes);
}

function countLaundryWeightKgClient(counts, itemTypes = laundryItemTypes()) {
  return Number(itemTypes.reduce((sum, item) => sum + (Number(counts?.[item.id] || 0) * Number(item.weightKg || 0)), 0).toFixed(2));
}

function formatLaundryKg(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "0,00 kg";
  const formatted = new Intl.NumberFormat("de-DE", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(num);
  return `${formatted} kg`;
}

function formatLaundryKgWithPrice(weightKg) {
  const amount = Math.max(0, Number(weightKg || 0));
  const pricePerKg = Math.max(0, Number(state.laundrySettings?.pricePerKg || 0));
  return `${formatLaundryKg(amount)} | ${formatMoney(amount * pricePerKg)}`;
}

function shiftLaundryDate(value, days) {
  const raw = clean(value);
  if (!raw) return "";
  const date = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(date.getTime())) return raw;
  date.setDate(date.getDate() + days);
  return formatDate(date);
}

function laundryReceiveDate(value) {
  return shiftLaundryDate(value, 2);
}

function getLatestLaundryDateForProperty(property, { excludeId = "" } = {}) {
  const normalizedProperty = normalizeLaundryPropertyClient(property);
  return state.laundryRecords
    .filter((row) => normalizeLaundryPropertyClient(row.property) === normalizedProperty && clean(row.id) !== clean(excludeId))
    .map((row) => clean(row.date))
    .filter(Boolean)
    .sort()
    .at(-1) || "";
}

function getRequiredNextLaundryDate(property, { excludeId = "" } = {}) {
  const latest = getLatestLaundryDateForProperty(property, { excludeId });
  return latest ? shiftLaundryDate(latest, 1) : "";
}

function validateLaundryDraftClient(draft, { isCreate = false } = {}) {
  if (!clean(draft?.date)) return "Sent date is required.";
  if (!clean(draft?.property)) return "Property is required.";
  const duplicate = state.laundryRecords.find((row) =>
    normalizeLaundryPropertyClient(row.property) === normalizeLaundryPropertyClient(draft.property) &&
    clean(row.date) === clean(draft.date) &&
    clean(row.id) !== clean(draft.id)
  );
  if (duplicate) return `A laundry record for ${normalizeLaundryPropertyClient(draft.property)} on ${clean(draft.date)} already exists.`;
  if (isCreate) {
    const expected = getRequiredNextLaundryDate(draft.property);
    if (expected && clean(draft.date) !== expected) {
      return `The next record for ${normalizeLaundryPropertyClient(draft.property)} must be created for ${expected}.`;
    }
  }
  return "";
}

function setLaundryStatus(text) {
  if (els.laundryStatus) els.laundryStatus.textContent = text;
}

function setLaundryDbStatus(text) {
  if (els.laundryDbStatus) els.laundryDbStatus.textContent = text;
}

function setLaundrySettingsStatus(text) {
  if (els.laundrySettingsStatus) els.laundrySettingsStatus.textContent = text;
}

async function loadLaundrySettings({ silent = false } = {}) {
  try {
    const result = await api("/api/laundry-settings");
    state.laundrySettings = normalizeLaundrySettingsClient(result?.settings);
    state.laundrySettingsLoaded = true;
    renderLaundrySettings();
    if (!silent) setLaundrySettingsStatus("Laundry configuration loaded.");
  } catch (e) {
    state.laundrySettings = clone(DEFAULT_LAUNDRY_SETTINGS);
    renderLaundrySettings();
    if (!silent) setLaundrySettingsStatus(`Using default laundry configuration (${e.message}).`);
  }
}

async function loadLaundryRecords({ silent = false } = {}) {
  try {
    const result = await api("/api/laundry");
    if (result?.settings) {
      state.laundrySettings = normalizeLaundrySettingsClient(result.settings);
      state.laundrySettingsLoaded = true;
    }
    state.laundryRecords = (Array.isArray(result?.rows) ? result.rows : []).map((row) => normalizeLaundryRecordClient(row, state.laundrySettings));
    state.laundryLoaded = true;
    renderLaundry();
    renderLaundrySettings();
    if (!silent) setLaundryDbStatus(`Loaded ${state.laundryRecords.length} laundry record${state.laundryRecords.length === 1 ? "" : "s"}.`);
  } catch (e) {
    state.laundryRecords = [];
    renderLaundry();
    if (!silent) {
      setLaundryDbStatus(`Failed to load laundry records: ${e.message}`);
      showToast(`Failed to load laundry records: ${e.message}`, "error");
    }
  }
}

function resetLaundryDraft() {
  state.laundrySelectedId = "";
  state.laundryDraft = emptyLaundryDraft();
}

function formatLaundryItemsSummary(counts) {
  const parts = laundryItemTypes()
    .map((item) => {
      const raw = counts?.[item.id];
      const filled = raw !== null && raw !== undefined && String(raw).trim() !== "";
      return { name: item.name, qty: filled ? Number(raw || 0) : null, filled };
    })
    .filter((item) => item.filled)
    .map((item) => `${item.name}: ${item.qty}`);
  return parts.length ? parts.join("\n") : "";
}

function formatLaundryColumnSummary(counts, weightKg) {
  const itemsText = formatLaundryItemsSummary(counts);
  return `${itemsText}\nKg: ${formatLaundryKg(weightKg)}`;
}

function formatLaundryListSummary(counts) {
  return formatLaundryItemsSummary(counts);
}

function laundryHasReceivedValues(record) {
  return laundryHasCompleteReceivedItemEntries(record);
}

function laundryHasReceivedItemEntries(record) {
  const itemTypes = laundryItemTypes();
  return itemTypes.some((item) => {
    const raw = record?.receivedItems?.[item.id];
    return raw !== null && raw !== undefined && String(raw).trim() !== "";
  });
}

function laundryHasCompleteReceivedItemEntries(record) {
  const itemTypes = laundryItemTypes();
  return itemTypes.every((item) => {
    const raw = record?.receivedItems?.[item.id];
    return raw !== null && raw !== undefined && String(raw).trim() !== "";
  });
}

function laundryHasMissingSentRecords() {
  const today = lisbonTodayIsoClient();
  return ["Hostel", "Cruz"].some((property) => {
    const latest = state.laundryRecords
      .filter((row) => clean(row?.property) === property)
      .map((row) => clean(row?.date))
      .filter(Boolean)
      .sort()
      .at(-1) || "";
    return !latest || latest < today;
  });
}

function describeLaundryMissingSentRecords() {
  const today = lisbonTodayIsoClient();
  const messages = [];
  ["Hostel", "Cruz"].forEach((property) => {
    const latest = state.laundryRecords
      .filter((row) => clean(row?.property) === property)
      .map((row) => clean(row?.date))
      .filter(Boolean)
      .sort()
      .at(-1) || "";
    if (latest && latest >= today) return;
    const firstMissing = latest ? shiftLaundryDate(latest, 1) : today;
    const dateLabel = firstMissing && firstMissing < today ? `${firstMissing} to ${today}` : (firstMissing || today);
    messages.push(`Missing records for ${dateLabel} for ${property}`);
  });
  return messages;
}

function laundryHasOverduePendingReceipts() {
  const today = lisbonTodayIsoClient();
  return state.laundryRecords.some((row) => {
    const receiveDate = clean(row?.receivedDate) || laundryReceiveDate(row?.date);
    if (!receiveDate || receiveDate > today) return false;
    return !laundryHasCompleteReceivedItemEntries(row);
  });
}

function describeLaundryPendingReceipts() {
  const today = lisbonTodayIsoClient();
  const seen = new Set();
  return state.laundryRecords
    .filter((row) => {
      const receiveDate = clean(row?.receivedDate) || laundryReceiveDate(row?.date);
      return !!receiveDate && receiveDate <= today && !laundryHasCompleteReceivedItemEntries(row);
    })
    .sort((a, b) => {
      const dateCompare = clean(a.receivedDate || laundryReceiveDate(a.date)).localeCompare(clean(b.receivedDate || laundryReceiveDate(b.date)));
      if (dateCompare !== 0) return dateCompare;
      return clean(a.property).localeCompare(clean(b.property));
    })
    .map((row) => {
      const receiveDate = clean(row?.receivedDate) || laundryReceiveDate(row?.date);
      return `Pending received quantities for ${receiveDate} for ${clean(row.property) || "-"}`;
    })
    .filter((message) => {
      if (seen.has(message)) return false;
      seen.add(message);
      return true;
    });
}

function buildLaundryDifferenceLines(record) {
  const itemTypes = laundryItemTypes();
  const receiveDate = clean(record?.receivedDate) || laundryReceiveDate(record?.date);
  const hasReceivedValues = laundryHasReceivedValues(record);
  const isReceiveDue = Boolean(receiveDate) && receiveDate <= lisbonTodayIsoClient();
  if (!hasReceivedValues) {
    const waitingText = receiveDate ? `Waiting for received quantities on ${receiveDate}.` : "Waiting for received quantities.";
    return {
      matchDate: receiveDate,
      hasReceivedValues,
      isReceiveDue,
      totalDiff: null,
      lines: [waitingText],
      entries: [{ text: waitingText, tone: isReceiveDue ? "negative" : "" }],
    };
  }
  const lines = [];
  const entries = [];
  let totalDiff = 0;
  itemTypes.forEach((item) => {
    const sent = Number(record.sentItems?.[item.id] || 0);
    const received = Number(record.receivedItems?.[item.id] || 0);
    const diff = received - sent;
    totalDiff += diff;
    if (diff !== 0) {
      const text = `${item.name}: ${sent} -> ${received} (${diff > 0 ? "+" : ""}${diff})`;
      lines.push(text);
      entries.push({ text, tone: diff > 0 ? "positive" : "negative" });
    }
  });
  const totalText = `Total counts difference: ${totalDiff > 0 ? "+" : ""}${totalDiff}`;
  lines.push(totalText);
  entries.push({ text: totalText, tone: totalDiff > 0 ? "positive" : totalDiff < 0 ? "negative" : "zero" });
  return {
    matchDate: receiveDate,
    hasReceivedValues,
    isReceiveDue,
    totalDiff,
    lines,
    entries,
  };
}

function renderLaundryItemInputs(container, counts, kind = "sent") {
  if (!container) return;
  container.innerHTML = laundryItemTypes().map((item) => `
    <label>
      <span>${escape(item.name)} <small>(${escape(String(Number(item.weightKg || 0).toFixed(2)))} kg)</small></span>
      <input data-laundry-count-kind="${escape(kind)}" data-laundry-item-id="${escape(item.id)}" type="number" min="0" step="1" value="${counts?.[item.id] === null || counts?.[item.id] === undefined || String(counts?.[item.id]).trim() === "" ? "" : escape(String(Number(counts?.[item.id] || 0)))}" />
    </label>
  `).join("");
}

function renderLaundryDraftComputedDetails(draft = state.laundryDraft || emptyLaundryDraft()) {
  const sentWeightKg = countLaundryWeightKgClient(draft.sentItems);
  const receivedComputedWeightKg = countLaundryWeightKgClient(draft.receivedItems);
  if (els.laundrySentWeight) els.laundrySentWeight.textContent = formatLaundryKgWithPrice(sentWeightKg);
  if (els.laundryReceivedComputedWeight) els.laundryReceivedComputedWeight.textContent = formatLaundryKgWithPrice(receivedComputedWeightKg);
  const diff = buildLaundryDifferenceLines(draft);
  if (els.laundryMatchDate) els.laundryMatchDate.textContent = diff.matchDate ? `Received date: ${diff.matchDate}` : "Received date pending.";
  if (els.laundryDifferenceSummary) {
    els.laundryDifferenceSummary.classList.toggle("empty", false);
    els.laundryDifferenceSummary.className = "laundry-diff-grid";
    els.laundryDifferenceSummary.innerHTML = (diff.entries || []).map((entry) => {
      const toneClass = entry.tone ? ` diff-${entry.tone}` : "";
      return `<article class="laundry-diff-pill${toneClass}">${escape(entry.text)}</article>`;
    }).join("");
  }
}

function renderLaundryDraft() {
  const focusTarget = document.activeElement?.matches?.("[data-laundry-count-kind][data-laundry-item-id], #laundry-received-weight, #laundry-notes") ? document.activeElement : null;
  const focusKind = focusTarget ? clean(focusTarget.dataset?.laundryCountKind) : "";
  const focusItemId = focusTarget ? clean(focusTarget.dataset?.laundryItemId) : "";
  const focusIsWeight = focusTarget === els.laundryReceivedWeight;
  const focusIsNotes = focusTarget === els.laundryNotes;
  const caretStart = focusTarget && typeof focusTarget.selectionStart === "number" ? focusTarget.selectionStart : null;
  const caretEnd = focusTarget && typeof focusTarget.selectionEnd === "number" ? focusTarget.selectionEnd : null;
  const draft = state.laundryDraft || emptyLaundryDraft();
  draft.receivedDate = laundryReceiveDate(draft.date);
  if (els.laundryProperty) els.laundryProperty.value = draft.property || "Hostel";
  if (els.laundryDate) els.laundryDate.value = draft.date || "";
  if (els.laundryReceiveDate) els.laundryReceiveDate.value = draft.receivedDate || "";
  if (els.laundryReceivedWeight) els.laundryReceivedWeight.value = draft.receivedWeightKg === "" || draft.receivedWeightKg === null || draft.receivedWeightKg === undefined ? "" : String(draft.receivedWeightKg);
  if (els.laundryNotes) els.laundryNotes.value = draft.notes || "";
  renderLaundryItemInputs(els.laundrySentItemsGrid, draft.sentItems, "sent");
  renderLaundryItemInputs(els.laundryReceivedItemsGrid, draft.receivedItems, "received");
  renderLaundryDraftComputedDetails(draft);
  let restoreTarget = null;
  if (focusKind && focusItemId) {
    restoreTarget = document.querySelector(`[data-laundry-count-kind="${focusKind}"][data-laundry-item-id="${focusItemId}"]`);
  } else if (focusIsWeight) {
    restoreTarget = els.laundryReceivedWeight;
  } else if (focusIsNotes) {
    restoreTarget = els.laundryNotes;
  }
  if (restoreTarget) {
    restoreTarget.focus();
    if (typeof caretStart === "number" && typeof restoreTarget.setSelectionRange === "function") {
      const nextStart = Math.min(caretStart, String(restoreTarget.value || "").length);
      const nextEnd = Math.min(caretEnd ?? caretStart, String(restoreTarget.value || "").length);
      restoreTarget.setSelectionRange(nextStart, nextEnd);
    }
  }
}

function openLaundryModal(recordId = "") {
  const needle = clean(recordId);
  if (needle) {
    const found = state.laundryRecords.find((item) => item.id === needle);
    if (found) {
      state.laundrySelectedId = found.id;
      state.laundryDraft = clone(found);
    }
  }
  if (!state.laundryDraft) resetLaundryDraft();
  els.laundryEditorModal.hidden = false;
  document.body.classList.add("modal-open");
  renderLaundryDraft();
  setLaundryStatus("");
}

function closeLaundryModal() {
  els.laundryEditorModal.hidden = true;
  document.body.classList.remove("modal-open");
  state.laundrySelectedId = "";
  state.laundryDraft = null;
}

function onLaundryFilterInput() {
  state.laundryFilters.property = clean(els.laundryFilterProperty?.value);
  state.laundryFilters.dateFrom = clean(els.laundryFilterDateFrom?.value);
  state.laundryFilters.dateTo = clean(els.laundryFilterDateTo?.value);
  state.laundryFilters.search = clean(els.laundryFilterSearch?.value);
  renderLaundry();
}

function onLaundryResumeFilterInput() {
  state.laundryResumeFilters.dateField = clean(els.laundryResumeDateField?.value) === "received" ? "received" : "sent";
  state.laundryResumeFilters.detail = !!els.laundryResumeDetail?.checked;
  state.laundryResumeFilters.property = clean(els.laundryResumeFilterProperty?.value);
  state.laundryResumeFilters.dateFrom = clean(els.laundryResumeFilterDateFrom?.value);
  state.laundryResumeFilters.dateTo = clean(els.laundryResumeFilterDateTo?.value);
  renderLaundry();
}

function onLaundryAnalysisFilterInput() {
  state.laundryResumeFilters.dateField = clean(els.laundryAnalysisDateField?.value) === "received" ? "received" : "sent";
  state.laundryResumeFilters.property = clean(els.laundryAnalysisFilterProperty?.value);
  state.laundryResumeFilters.dateFrom = clean(els.laundryAnalysisFilterDateFrom?.value);
  state.laundryResumeFilters.dateTo = clean(els.laundryAnalysisFilterDateTo?.value);
  renderLaundry();
}

function setLaundryScreen(screen) {
  state.laundryScreen = screen === "resume" || screen === "analysis" ? screen : "list";
  renderLaundry();
}

function renderLaundryScreenTabs() {
  const isList = state.laundryScreen === "list";
  const isResume = state.laundryScreen === "resume";
  const isAnalysis = state.laundryScreen === "analysis";
  els.laundryTabList?.classList.toggle("active-tab", isList);
  els.laundryTabList?.classList.toggle("ghost", !isList);
  els.laundryTabResume?.classList.toggle("active-tab", isResume);
  els.laundryTabResume?.classList.toggle("ghost", !isResume);
  els.laundryTabAnalysis?.classList.toggle("active-tab", isAnalysis);
  els.laundryTabAnalysis?.classList.toggle("ghost", !isAnalysis);
  if (els.laundryPanelList) els.laundryPanelList.hidden = !isList;
  if (els.laundryPanelResume) els.laundryPanelResume.hidden = !isResume;
  if (els.laundryPanelAnalysis) els.laundryPanelAnalysis.hidden = !isAnalysis;
}

function onLaundryDraftInput(event) {
  if (!state.laundryDraft) return;
  const target = event.target;
  if (target === els.laundryProperty) {
    state.laundryDraft.property = normalizeLaundryPropertyClient(target.value);
    if (!state.laundrySelectedId) {
      state.laundryDraft.date = getRequiredNextLaundryDate(state.laundryDraft.property) || formatDate(new Date());
    }
  }
  if (target === els.laundryDate) state.laundryDraft.date = clean(target.value);
  if (target === els.laundryReceivedWeight) {
    const normalized = normalizeNumber(target.value);
    state.laundryDraft.receivedWeightKg = normalized == null ? "" : Math.max(0, Number(normalized || 0));
    return;
  }
  if (target === els.laundryNotes) {
    state.laundryDraft.notes = clean(target.value);
    return;
  }
  renderLaundryDraft();
}

function onLaundryDraftGridInput(event) {
  if (!state.laundryDraft) return;
  const kind = clean(event.target.dataset.laundryCountKind);
  const itemId = clean(event.target.dataset.laundryItemId);
  if (!kind || !itemId) return;
  const normalized = normalizeNumber(event.target.value);
  const amount = normalized == null ? null : Math.max(0, Math.round(Number(normalized || 0)));
  if (kind === "sent") state.laundryDraft.sentItems[itemId] = amount;
  if (kind === "received") state.laundryDraft.receivedItems[itemId] = amount;
  renderLaundryDraftComputedDetails(state.laundryDraft);
}

function getFilteredLaundryRecords() {
  const filters = state.laundryFilters || {};
  const property = clean(filters.property);
  const dateFrom = clean(filters.dateFrom);
  const dateTo = clean(filters.dateTo);
  const search = clean(filters.search).toLowerCase();
  return [...state.laundryRecords]
    .filter((row) => !property || clean(row.property) === property)
    .filter((row) => !dateFrom || clean(row.date) >= dateFrom)
    .filter((row) => !dateTo || clean(row.date) <= dateTo)
    .filter((row) => {
      if (!search) return true;
    const haystack = [row.notes, formatLaundryItemsSummary(row.sentItems), formatLaundryItemsSummary(row.receivedItems)]
        .join("\n")
        .toLowerCase();
      return haystack.includes(search);
    })
    .sort((a, b) => {
      const dateCompare = clean(b.date).localeCompare(clean(a.date));
      if (dateCompare !== 0) return dateCompare;
      if (a.property === b.property) return 0;
      return a.property === "Hostel" ? -1 : 1;
    });
}

function getLaundryResumeColumnIds() {
  const itemTypes = laundryItemTypes();
  const findId = (needle) => itemTypes.find((item) => clean(item?.id) === needle || clean(item?.name).toLowerCase() === needle.replace("-", " "))?.id || needle;
  return {
    singleBaixo: findId("single-baixo"),
    singleCima: findId("single-cima"),
    casalBaixo: findId("casal-baixo"),
    casalCima: findId("casal-cima"),
  };
}

function formatLaundryMonthLabel(monthKey) {
  const raw = clean(monthKey);
  if (!/^\d{4}-\d{2}$/.test(raw)) return raw || "-";
  const dt = new Date(`${raw}-01T00:00:00`);
  if (Number.isNaN(dt.getTime())) return raw;
  return new Intl.DateTimeFormat("en-GB", { month: "short", year: "numeric", timeZone: "Europe/Lisbon" }).format(dt);
}

function getLaundryResumeRows() {
  const filters = state.laundryResumeFilters || {};
  const dateField = clean(filters.dateField) === "received" ? "received" : "sent";
  const detail = !!filters.detail;
  const property = clean(filters.property);
  const dateFrom = clean(filters.dateFrom);
  const dateTo = clean(filters.dateTo);
  const ids = getLaundryResumeColumnIds();
  const buckets = new Map();
  state.laundryRecords.forEach((row) => {
    if (!laundryHasCompleteReceivedItemEntries(row)) return;
    const basisDate = clean(dateField === "received" ? row.receivedDate : row.date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(basisDate)) return;
    if (property && clean(row.property) !== property) return;
    if (dateFrom && basisDate < dateFrom) return;
    if (dateTo && basisDate > dateTo) return;
    const monthKey = basisDate.slice(0, 7);
    const current = buckets.get(monthKey) || {
      monthKey,
      records: [],
      difSingleBaixo: 0,
      difSingleCima: 0,
      difCasalBaixo: 0,
      difCasalCima: 0,
      totalDiff: 0,
      receivedWeightKg: 0,
    };
    const sentItems = row.sentItems || {};
    const receivedItems = row.receivedItems || {};
    const difSingleBaixo = Number(receivedItems?.[ids.singleBaixo] || 0) - Number(sentItems?.[ids.singleBaixo] || 0);
    const difSingleCima = Number(receivedItems?.[ids.singleCima] || 0) - Number(sentItems?.[ids.singleCima] || 0);
    const difCasalBaixo = Number(receivedItems?.[ids.casalBaixo] || 0) - Number(sentItems?.[ids.casalBaixo] || 0);
    const difCasalCima = Number(receivedItems?.[ids.casalCima] || 0) - Number(sentItems?.[ids.casalCima] || 0);
    const totalDiff = difSingleBaixo + difSingleCima + difCasalBaixo + difCasalCima;
    const receivedWeightKg = countLaundryWeightKgClient(receivedItems);
    current.difSingleBaixo += difSingleBaixo;
    current.difSingleCima += difSingleCima;
    current.difCasalBaixo += difCasalBaixo;
    current.difCasalCima += difCasalCima;
    current.totalDiff += totalDiff;
    current.receivedWeightKg += receivedWeightKg;
    if (detail) {
      const existingRecord = (current.records || []).find((item) => clean(item.date) === basisDate);
      if (existingRecord) {
        existingRecord.difSingleBaixo += difSingleBaixo;
        existingRecord.difSingleCima += difSingleCima;
        existingRecord.difCasalBaixo += difCasalBaixo;
        existingRecord.difCasalCima += difCasalCima;
        existingRecord.totalDiff += totalDiff;
        existingRecord.receivedWeightKg = Number((existingRecord.receivedWeightKg + receivedWeightKg).toFixed(2));
      } else {
        current.records.push({
          date: basisDate,
          difSingleBaixo,
          difSingleCima,
          difCasalBaixo,
          difCasalCima,
          totalDiff,
          receivedWeightKg,
        });
      }
    }
    buckets.set(monthKey, current);
  });
  return [...buckets.values()]
    .map((row) => ({
      ...row,
      receivedWeightKg: Number(row.receivedWeightKg.toFixed(2)),
      records: detail
        ? [...row.records].sort((a, b) => clean(b.date).localeCompare(clean(a.date)))
        : [],
    }))
    .sort((a, b) => clean(b.monthKey).localeCompare(clean(a.monthKey)));
}

function getLaundryCompletedAnalysisEntries() {
  const filters = state.laundryResumeFilters || {};
  const dateField = clean(filters.dateField) === "received" ? "received" : "sent";
  const property = clean(filters.property);
  const dateFrom = clean(filters.dateFrom);
  const dateTo = clean(filters.dateTo);
  const itemTypes = laundryItemTypes();
  const grouped = new Map();
  state.laundryRecords.forEach((row) => {
    if (!laundryHasCompleteReceivedItemEntries(row)) return;
    const basisDate = clean(dateField === "received" ? row.receivedDate : row.date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(basisDate)) return;
    if (property && clean(row.property) !== property) return;
    if (dateFrom && basisDate < dateFrom) return;
    if (dateTo && basisDate > dateTo) return;
    const sentItems = row.sentItems || {};
    const receivedItems = row.receivedItems || {};
    const current = grouped.get(basisDate) || {
      date: basisDate,
      diffs: Object.fromEntries(itemTypes.map((item) => [item.id, 0])),
      total: 0,
    };
    itemTypes.forEach((item) => {
      const diff = Number(receivedItems?.[item.id] || 0) - Number(sentItems?.[item.id] || 0);
      current.diffs[item.id] += diff;
      current.total += diff;
    });
    grouped.set(basisDate, current);
  });
  return [...grouped.values()].sort((a, b) => clean(a.date).localeCompare(clean(b.date)));
}

const LAUNDRY_ANALYSIS_COLORS = ["#2563eb", "#ef4444", "#16a34a", "#f59e0b", "#7c3aed", "#0891b2", "#be185d", "#4b5563"];

function buildLaundryAnalysisSeries() {
  const itemTypes = laundryItemTypes();
  const entries = getLaundryCompletedAnalysisEntries();
  const labels = entries.map((entry) => entry.date);
  const running = Object.fromEntries(itemTypes.map((item) => [item.id, 0]));
  let runningTotal = 0;
  const series = itemTypes.map((item, index) => ({
    key: item.id,
    label: item.name,
    color: LAUNDRY_ANALYSIS_COLORS[index % LAUNDRY_ANALYSIS_COLORS.length],
    points: [],
  }));
  const totalSeries = {
    key: "total",
    label: "Total",
    color: "#111827",
    points: [],
  };
  entries.forEach((entry, index) => {
    itemTypes.forEach((item, itemIndex) => {
      running[item.id] += Number(entry.diffs?.[item.id] || 0);
      series[itemIndex].points.push({
        index,
        value: running[item.id],
        delta: Number(entry.diffs?.[item.id] || 0),
        date: entry.date,
      });
    });
    runningTotal += Number(entry.total || 0);
    totalSeries.points.push({
      index,
      value: runningTotal,
      delta: Number(entry.total || 0),
      date: entry.date,
    });
  });
  return {
    labels,
    entries,
    series: [...series, totalSeries].filter((row) => row.points.length > 0),
  };
}

function laundryAnalysisScale(series) {
  const values = series.flatMap((item) => item.points.map((point) => point.value)).filter((value) => Number.isFinite(value));
  if (!values.length) return { min: -1, max: 1 };
  const minRaw = Math.min(...values, 0);
  const maxRaw = Math.max(...values, 0);
  let min = minRaw;
  let max = maxRaw;
  if (min === max) {
    min -= 1;
    max += 1;
  }
  const padding = Math.max(1, Math.ceil((max - min) * 0.08));
  return { min: min - padding, max: max + padding };
}

function laundryAnalysisTicks(scale) {
  const span = scale.max - scale.min;
  const rough = span / 5;
  const magnitude = Math.pow(10, Math.floor(Math.log10(Math.max(rough, 1))));
  const normalized = rough / magnitude;
  const stepBase = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  const step = stepBase * magnitude;
  const start = Math.floor(scale.min / step) * step;
  const ticks = [];
  for (let value = start; value <= scale.max + step * 0.5; value += step) {
    ticks.push(Number(value.toFixed(6)));
  }
  return ticks;
}

function renderLaundryAnalysisChart() {
  if (!els.laundryAnalysisChart || !els.laundryAnalysisLegend || !els.laundryAnalysisStatus) return;
  const analysis = buildLaundryAnalysisSeries();
  els.laundryAnalysisChart.innerHTML = "";
  els.laundryAnalysisLegend.innerHTML = "";
  if (!analysis.labels.length || !analysis.series.length) {
    els.laundryAnalysisStatus.textContent = "No completed laundry records in the current filters";
    els.laundryAnalysisChart.innerHTML = '<div class="empty">No completed laundry records available for this chart.</div>';
    return;
  }
  els.laundryAnalysisStatus.textContent = `${analysis.labels.length} date${analysis.labels.length === 1 ? "" : "s"} shown`;
  const width = 920;
  const height = 330;
  const margin = { top: 22, right: 24, bottom: 58, left: 56 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const scale = laundryAnalysisScale(analysis.series);
  const x = (index) => margin.left + (analysis.labels.length === 1 ? plotWidth / 2 : (index / (analysis.labels.length - 1)) * plotWidth);
  const y = (value) => margin.top + ((scale.max - value) / (scale.max - scale.min)) * plotHeight;
  const labelStep = Math.max(1, Math.ceil(analysis.labels.length / 6));
  const labelIndexes = analysis.labels
    .map((label, index) => ({ label, index }))
    .filter((item, index) => index === 0 || index === analysis.labels.length - 1 || index % labelStep === 0);
  const gridLines = laundryAnalysisTicks(scale).map((value) => {
    const yy = y(value);
    return `<line x1="${margin.left}" y1="${yy}" x2="${width - margin.right}" y2="${yy}" class="analysis-grid-line" />
      <text x="${margin.left - 10}" y="${yy + 4}" text-anchor="end" class="analysis-axis-label">${escape(String(value))}</text>`;
  }).join("");
  const zeroLine = scale.min <= 0 && scale.max >= 0
    ? `<line x1="${margin.left}" y1="${y(0)}" x2="${width - margin.right}" y2="${y(0)}" class="analysis-axis-line" />`
    : "";
  const labels = labelIndexes.map(({ label, index }) => {
    const xx = x(index);
    return `<text x="${xx}" y="${height - 24}" text-anchor="end" transform="rotate(-35 ${xx} ${height - 24})" class="analysis-axis-label">${escape(label)}</text>`;
  }).join("");
  const seriesSvg = analysis.series.map((series) => {
    const coordinates = series.points.map((point) => `${x(point.index).toFixed(1)},${y(point.value).toFixed(1)}`).join(" ");
    const circles = series.points.map((point) => {
      const tooltip = `${series.label} | ${point.date} | Accumulated ${point.value} | Change ${point.delta >= 0 ? "+" : ""}${point.delta}`;
      return `<circle class="analysis-point laundry-analysis-point" cx="${x(point.index).toFixed(1)}" cy="${y(point.value).toFixed(1)}" r="${series.key === "total" ? 4 : 3}" fill="${series.color}" tabindex="0" data-tooltip="${escape(tooltip)}">
        <title>${escape(tooltip)}</title>
      </circle>`;
    }).join("");
    return `<polyline points="${coordinates}" fill="none" stroke="${series.color}" stroke-width="${series.key === "total" ? 3.2 : 2}" stroke-linecap="round" stroke-linejoin="round" opacity="${series.key === "total" ? "1" : "0.88"}" />${circles}`;
  }).join("");
  els.laundryAnalysisChart.innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Accumulated laundry differences by date">
    <rect x="0" y="0" width="${width}" height="${height}" rx="16" class="analysis-chart-bg" />
    ${gridLines}
    ${zeroLine}
    <line x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${height - margin.bottom}" class="analysis-axis-line" />
    <line x1="${margin.left}" y1="${height - margin.bottom}" x2="${width - margin.right}" y2="${height - margin.bottom}" class="analysis-axis-line" />
    ${labels}
    ${seriesSvg}
  </svg><div class="analysis-tooltip" role="status" aria-live="polite"></div>`;
  els.laundryAnalysisLegend.innerHTML = analysis.series.map((series) =>
    `<span class="analysis-legend-item"><span class="analysis-legend-swatch" style="background:${escape(series.color)}"></span>${escape(series.label)}</span>`
  ).join("");
  bindAnalysisTooltip(els.laundryAnalysisChart);
}

function bindAnalysisTooltip(container) {
  if (!container) return;
  const tooltip = container.querySelector(".analysis-tooltip");
  if (!tooltip) return;
  const hideTooltip = () => {
    tooltip.classList.remove("visible");
    tooltip.textContent = "";
  };
  container.querySelectorAll(".analysis-point").forEach((point) => {
    const showTooltip = () => {
      const chartRect = container.getBoundingClientRect();
      const pointRect = point.getBoundingClientRect();
      tooltip.textContent = point.dataset.tooltip || "";
      tooltip.style.left = `${pointRect.left - chartRect.left + pointRect.width / 2}px`;
      tooltip.style.top = `${pointRect.top - chartRect.top}px`;
      tooltip.classList.add("visible");
    };
    point.addEventListener("mouseenter", showTooltip);
    point.addEventListener("focus", showTooltip);
    point.addEventListener("mouseleave", hideTooltip);
    point.addEventListener("blur", hideTooltip);
  });
}

function renderLaundryResume() {
  if (!els.laundryResumeBody || !els.laundryResumeCount) return;
  els.laundryResumeDateField.value = clean(state.laundryResumeFilters.dateField) === "received" ? "received" : "sent";
  els.laundryResumeDetail.checked = !!state.laundryResumeFilters.detail;
  els.laundryResumeFilterProperty.value = clean(state.laundryResumeFilters.property);
  els.laundryResumeFilterDateFrom.value = clean(state.laundryResumeFilters.dateFrom);
  els.laundryResumeFilterDateTo.value = clean(state.laundryResumeFilters.dateTo);
  const rows = getLaundryResumeRows();
  const detail = !!state.laundryResumeFilters.detail;
  const detailCount = rows.reduce((sum, row) => sum + (row.records?.length || 0), 0);
  const pricePerKg = Math.max(0, Number(state.laundrySettings?.pricePerKg || 0));
  els.laundryResumeCount.textContent = detail ? `${rows.length} month${rows.length === 1 ? "" : "s"} / ${detailCount} record${detailCount === 1 ? "" : "s"}` : `${rows.length} month${rows.length === 1 ? "" : "s"}`;
  els.laundryResumeBody.innerHTML = "";
  if (!rows.length) {
    els.laundryResumeBody.innerHTML = '<tr><td colspan="8" class="empty">No laundry resume data found.</td></tr>';
    return;
  }
  const overall = rows.reduce((acc, row) => {
    acc.difSingleBaixo += Number(row.difSingleBaixo || 0);
    acc.difSingleCima += Number(row.difSingleCima || 0);
    acc.difCasalBaixo += Number(row.difCasalBaixo || 0);
    acc.difCasalCima += Number(row.difCasalCima || 0);
    acc.totalDiff += Number(row.totalDiff || 0);
    acc.receivedWeightKg += Number(row.receivedWeightKg || 0);
    return acc;
  }, {
    difSingleBaixo: 0,
    difSingleCima: 0,
    difCasalBaixo: 0,
    difCasalCima: 0,
    totalDiff: 0,
    receivedWeightKg: 0,
  });
  rows.forEach((row) => {
    if (detail) {
      (row.records || []).forEach((record) => {
        const detailTr = document.createElement("tr");
        if (record.totalDiff > 0) detailTr.classList.add("laundry-row-positive");
        else if (record.totalDiff < 0) detailTr.classList.add("laundry-row-negative");
        else detailTr.classList.add("laundry-row-zero");
        detailTr.innerHTML = `<td>${escape(record.date)}</td>
          <td>${escape(String(record.difSingleBaixo))}</td>
          <td>${escape(String(record.difSingleCima))}</td>
          <td>${escape(String(record.difCasalBaixo))}</td>
          <td>${escape(String(record.difCasalCima))}</td>
          <td>${escape(String(record.totalDiff))}</td>
          <td>${escape(formatLaundryKg(record.receivedWeightKg))}</td>
          <td>${escape(formatMoney(record.receivedWeightKg * pricePerKg))}</td>`;
        els.laundryResumeBody.appendChild(detailTr);
      });
    }
    const tr = document.createElement("tr");
    tr.classList.add("laundry-resume-total-row");
    if (row.totalDiff > 0) tr.classList.add("laundry-row-positive");
    else if (row.totalDiff < 0) tr.classList.add("laundry-row-negative");
    else tr.classList.add("laundry-row-zero");
    tr.innerHTML = `<td>${escape(detail ? `Total ${formatLaundryMonthLabel(row.monthKey)}` : formatLaundryMonthLabel(row.monthKey))}</td>
      <td>${escape(String(row.difSingleBaixo))}</td>
      <td>${escape(String(row.difSingleCima))}</td>
      <td>${escape(String(row.difCasalBaixo))}</td>
      <td>${escape(String(row.difCasalCima))}</td>
      <td>${escape(String(row.totalDiff))}</td>
      <td>${escape(formatLaundryKg(row.receivedWeightKg))}</td>
      <td>${escape(formatMoney(row.receivedWeightKg * pricePerKg))}</td>`;
    els.laundryResumeBody.appendChild(tr);
  });
  const overallTr = document.createElement("tr");
  overallTr.classList.add("laundry-resume-total-row");
  if (overall.totalDiff > 0) overallTr.classList.add("laundry-row-positive");
  else if (overall.totalDiff < 0) overallTr.classList.add("laundry-row-negative");
  else overallTr.classList.add("laundry-row-zero");
  overallTr.innerHTML = `<td>Overall Total</td>
    <td>${escape(String(overall.difSingleBaixo))}</td>
    <td>${escape(String(overall.difSingleCima))}</td>
    <td>${escape(String(overall.difCasalBaixo))}</td>
    <td>${escape(String(overall.difCasalCima))}</td>
    <td>${escape(String(overall.totalDiff))}</td>
    <td>${escape(formatLaundryKg(Number(overall.receivedWeightKg.toFixed(2))))}</td>
    <td>${escape(formatMoney(overall.receivedWeightKg * pricePerKg))}</td>`;
  els.laundryResumeBody.appendChild(overallTr);
}

function laundryRawCountValue(counts, itemId) {
  const raw = counts?.[itemId];
  if (raw === null || raw === undefined || String(raw).trim() === "") return "";
  return String(Number(raw || 0));
}

function buildLaundryExportRows() {
  const itemTypes = laundryItemTypes();
  const pricePerKg = Math.max(0, Number(state.laundrySettings?.pricePerKg || 0));
  return getFilteredLaundryRecords().map((row) => {
    const sentWeightKg = countLaundryWeightKgClient(row.sentItems || {});
    const hasReceivedComplete = laundryHasCompleteReceivedItemEntries(row);
    const receivedWeightKg = hasReceivedComplete ? countLaundryWeightKgClient(row.receivedItems || {}) : null;
    const receivedKgManual = row.receivedWeightKg === null || row.receivedWeightKg === undefined || String(row.receivedWeightKg).trim() === "" ? "" : Number(row.receivedWeightKg || 0);
    const sentCounts = Object.fromEntries(itemTypes.map((item) => [item.id, laundryRawCountValue(row.sentItems, item.id)]));
    const receivedCounts = Object.fromEntries(itemTypes.map((item) => [item.id, laundryRawCountValue(row.receivedItems, item.id)]));
    const diffs = Object.fromEntries(itemTypes.map((item) => {
      const sentRaw = row.sentItems?.[item.id];
      const receivedRaw = row.receivedItems?.[item.id];
      const sentFilled = sentRaw !== null && sentRaw !== undefined && String(sentRaw).trim() !== "";
      const receivedFilled = receivedRaw !== null && receivedRaw !== undefined && String(receivedRaw).trim() !== "";
      const value = receivedFilled ? Number(receivedRaw || 0) - Number(sentRaw || 0) : "";
      return [item.id, sentFilled || receivedFilled ? value : ""];
    }));
    const totalDifference = hasReceivedComplete
      ? itemTypes.reduce((sum, item) => sum + Number(diffs[item.id] || 0), 0)
      : "";
    const price = receivedWeightKg == null ? "" : formatMoney(receivedWeightKg * pricePerKg);
    return {
      sentDate: row.date,
      property: row.property,
      sentCounts,
      sentWeightKg,
      receivedDate: row.receivedDate || laundryReceiveDate(row.date),
      receivedCounts,
      calculatedReceivedWeightKg: receivedWeightKg,
      receivedKg: receivedKgManual,
      diffs,
      totalDifference,
      price,
      notes: row.notes || "",
    };
  });
}

function exportLaundryToExcel() {
  const itemTypes = laundryItemTypes();
  const rows = buildLaundryExportRows();
  if (!rows.length) {
    showToast("No laundry records to export.", "error");
    return;
  }
  const headers = [
    "Sent Date",
    "Property",
    ...itemTypes.map((item) => `Sent ${item.name}`),
    "Calculated sent Kg",
    "Received Date",
    ...itemTypes.map((item) => `Received ${item.name}`),
    "Calculated received Kg",
    "Received Kg",
    ...itemTypes.map((item) => `Dif ${item.name}`),
    "Total Difference",
    "Price",
    "Notes",
  ];
  const bodyRows = rows.map((row) => [
    row.sentDate,
    row.property,
    ...itemTypes.map((item) => row.sentCounts[item.id]),
    Number(row.sentWeightKg || 0).toFixed(2),
    row.receivedDate,
    ...itemTypes.map((item) => row.receivedCounts[item.id]),
    row.calculatedReceivedWeightKg == null ? "" : Number(row.calculatedReceivedWeightKg).toFixed(2),
    row.receivedKg === "" ? "" : Number(row.receivedKg).toFixed(2),
    ...itemTypes.map((item) => row.diffs[item.id]),
    row.totalDifference,
    row.price,
    row.notes,
  ]);
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><table border="1"><thead><tr>${headers.map((cell) => `<th>${escape(cell)}</th>`).join("")}</tr></thead><tbody>${bodyRows.map((cells) => `<tr>${cells.map((cell) => `<td>${escape(cell === null || cell === undefined ? "" : String(cell))}</td>`).join("")}</tr>`).join("")}</tbody></table></body></html>`;
  downloadBlob(`laundry_records_${formatDate(new Date())}.xls`, html, "application/vnd.ms-excel;charset=utf-8;");
  showToast(`Exported ${rows.length} laundry record${rows.length === 1 ? "" : "s"} to Excel.`, "success");
}

function renderLaundry() {
  if (!canApp("laundry")) {
    if (els.laundryCount) els.laundryCount.textContent = "0 records";
    if (els.laundryMissingWarning) els.laundryMissingWarning.hidden = true;
    if (els.laundryRows) els.laundryRows.innerHTML = '<tr><td colspan="7" class="empty">Your profile has no access to Laundry Control.</td></tr>';
    if (els.laundryMobileCards) {
      els.laundryMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Laundry Control.</div>';
    }
    return;
  }
  renderLaundryScreenTabs();
  if (state.laundryScreen === "resume") {
    renderLaundryResume();
    return;
  }
  if (state.laundryScreen === "analysis") {
    if (els.laundryAnalysisDateField) els.laundryAnalysisDateField.value = clean(state.laundryResumeFilters.dateField) === "received" ? "received" : "sent";
    if (els.laundryAnalysisFilterProperty) els.laundryAnalysisFilterProperty.value = clean(state.laundryResumeFilters.property);
    if (els.laundryAnalysisFilterDateFrom) els.laundryAnalysisFilterDateFrom.value = clean(state.laundryResumeFilters.dateFrom);
    if (els.laundryAnalysisFilterDateTo) els.laundryAnalysisFilterDateTo.value = clean(state.laundryResumeFilters.dateTo);
    renderLaundryAnalysisChart();
    return;
  }
  const rows = getFilteredLaundryRecords();
  const warningMessages = [
    ...describeLaundryMissingSentRecords(),
    ...describeLaundryPendingReceipts(),
  ];
  if (els.laundryCount) els.laundryCount.textContent = `${rows.length} record${rows.length === 1 ? "" : "s"}`;
  if (els.laundryMissingWarning) {
    els.laundryMissingWarning.hidden = warningMessages.length === 0;
    els.laundryMissingWarning.innerHTML = warningMessages.map((line) => escape(line)).join("<br>");
  }
  if (els.laundryFilterProperty) els.laundryFilterProperty.value = state.laundryFilters.property;
  if (els.laundryFilterDateFrom) els.laundryFilterDateFrom.value = state.laundryFilters.dateFrom;
  if (els.laundryFilterDateTo) els.laundryFilterDateTo.value = state.laundryFilters.dateTo;
  if (els.laundryFilterSearch) els.laundryFilterSearch.value = state.laundryFilters.search;
  if (els.laundryRows) els.laundryRows.innerHTML = "";
  renderLaundryMobileCards(rows);
  if (!rows.length) {
    els.laundryRows.innerHTML = '<tr><td colspan="7" class="empty">No laundry records found.</td></tr>';
    return;
  }
  rows.forEach((row) => {
    const diff = buildLaundryDifferenceLines(row);
    const receiveDate = clean(row.receivedDate) || laundryReceiveDate(row.date) || "";
    const hasReceivedValues = diff.hasReceivedValues;
    const receiveDateNeedsFill = Boolean(receiveDate) && receiveDate <= lisbonTodayIsoClient() && !hasReceivedValues;
    const receivedSummary = hasReceivedValues
      ? escape(formatLaundryListSummary(row.receivedItems)).replace(/\n/g, "<br>")
      : "";
    const receivedDateCell = receiveDate
      ? `${escape(receiveDate)}${receiveDateNeedsFill ? '<br><span class="warning-text">please fill</span>' : ""}`
      : "-";
    const diffHtml = diff.lines
      .map((line) => (!hasReceivedValues && diff.isReceiveDue) ? `<span class="warning-text">${escape(line)}</span>` : escape(line))
      .join("<br>");
    const tr = document.createElement("tr");
    tr.className = "clickable-row";
    if (hasReceivedValues) {
      if (diff.totalDiff > 0) tr.classList.add("laundry-row-positive");
      else if (diff.totalDiff < 0) tr.classList.add("laundry-row-negative");
      else tr.classList.add("laundry-row-zero");
    }
    tr.dataset.laundryId = row.id;
    tr.innerHTML = `<td>${escape(row.date)}</td>
      <td>${escape(row.property)}</td>
      <td>${escape(formatLaundryListSummary(row.sentItems) || "-").replace(/\n/g, "<br>")}</td>
      <td>${receivedDateCell}</td>
      <td>${receivedSummary || "-"}</td>
      <td>${diffHtml}</td>
      <td>${escape(row.notes || "-")}</td>`;
    els.laundryRows.appendChild(tr);
  });
}

function renderLaundryMobileCards(rows) {
  if (!els.laundryMobileCards) return;
  const list = els.laundryMobileCards;
  list.innerHTML = "";
  if (!rows.length) {
    list.innerHTML = '<div class="services-mobile-empty">No laundry records found.</div>';
    return;
  }
  rows.forEach((row) => {
    list.appendChild(buildLaundryMobileCard(row));
  });
}

function buildLaundryMobileCard(row) {
  const diff = buildLaundryDifferenceLines(row);
  const receiveDate = clean(row.receivedDate) || laundryReceiveDate(row.date) || "";
  const hasReceivedValues = diff.hasReceivedValues;
  const receiveDateNeedsFill = Boolean(receiveDate) && receiveDate <= lisbonTodayIsoClient() && !hasReceivedValues;
  const receivedSummary = hasReceivedValues ? formatLaundryListSummary(row.receivedItems) : "";
  const card = document.createElement("article");
  card.className = "laundry-mobile-card";
  card.dataset.laundryId = row.id;
  if (hasReceivedValues) {
    if (diff.totalDiff > 0) card.classList.add("laundry-row-positive");
    else if (diff.totalDiff < 0) card.classList.add("laundry-row-negative");
    else card.classList.add("laundry-row-zero");
  }
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(row.property)}</div>
        <div class="communication-mobile-meta">Sent ${escape(row.date)}</div>
      </div>
      <div class="laundry-mobile-date-block">
        <small>Received</small>
        <strong>${receiveDate ? escape(receiveDate) : "-"}</strong>
        ${receiveDateNeedsFill ? '<span class="warning-text">please fill</span>' : ""}
      </div>
    </div>
    <div class="communication-mobile-grid">
      <div class="communication-mobile-field">
        <small>Sent</small>
        <div class="communication-mobile-message">${escape(formatLaundryListSummary(row.sentItems) || "-").replace(/\n/g, "<br>")}</div>
      </div>
      <div class="communication-mobile-field">
        <small>Received</small>
        <div class="communication-mobile-message">${receivedSummary ? escape(receivedSummary).replace(/\n/g, "<br>") : "-"}</div>
      </div>
      <div class="communication-mobile-field communication-mobile-field-full">
        <small>Difference</small>
        <div class="communication-mobile-message">${diff.lines.map((line) => (!hasReceivedValues && diff.isReceiveDue) ? `<span class="warning-text">${escape(line)}</span>` : escape(line)).join("<br>")}</div>
      </div>
      <div class="communication-mobile-field communication-mobile-field-full">
        <small>Notes</small>
        <div class="communication-mobile-message">${escape(row.notes || "-")}</div>
      </div>
    </div>`;
  return card;
}

function renderLaundrySettings() {
  if (!canSettings("laundry")) return;
  const settings = state.laundrySettings || clone(DEFAULT_LAUNDRY_SETTINGS);
  if (els.laundryPricePerKg) els.laundryPricePerKg.value = settings.pricePerKg || "";
  if (els.laundryEmailRecipients) els.laundryEmailRecipients.value = (settings.emailRecipients || []).join("\n");
  if (els.laundryEmailEnabled) els.laundryEmailEnabled.checked = !!settings.emailEnabled;
  if (els.laundryEmailTime) els.laundryEmailTime.value = settings.emailTime || "00:00";
  if (els.laundryManagementEmailRecipients) els.laundryManagementEmailRecipients.value = (settings.managementEmailRecipients || []).join("\n");
  if (els.laundryManagementEmailEnabled) els.laundryManagementEmailEnabled.checked = !!settings.managementEmailEnabled;
  if (els.laundryManagementEmailTime) els.laundryManagementEmailTime.value = settings.managementEmailTime || "00:00";
  if (els.laundryItemTypesBody) {
    els.laundryItemTypesBody.innerHTML = "";
    settings.itemTypes.forEach((item, index) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td><input data-laundry-setting-field="name" data-index="${index}" value="${escape(item.name)}" /></td>
        <td><input data-laundry-setting-field="weightKg" data-index="${index}" type="number" min="0" step="0.01" value="${escape(String(item.weightKg || 0))}" /></td>
        <td class="row-actions"><button type="button" class="ghost" data-action="remove-laundry-type" data-index="${index}">Remove</button></td>`;
      els.laundryItemTypesBody.appendChild(tr);
    });
  }
}

function onLaundrySettingsInput() {
  state.laundrySettings.pricePerKg = Math.max(0, Number(normalizeNumber(els.laundryPricePerKg?.value) || 0));
  state.laundrySettings.emailRecipients = String(els.laundryEmailRecipients?.value || "")
    .split(/[\n,;]/)
    .map((item) => clean(item).toLowerCase())
    .filter(Boolean)
    .filter((item, index, items) => items.indexOf(item) === index);
  state.laundrySettings.emailEnabled = !!els.laundryEmailEnabled?.checked;
  state.laundrySettings.emailTime = normalizeTimeInput(els.laundryEmailTime?.value);
  state.laundrySettings.managementEmailRecipients = String(els.laundryManagementEmailRecipients?.value || "")
    .split(/[\n,;]/)
    .map((item) => clean(item).toLowerCase())
    .filter(Boolean)
    .filter((item, index, items) => items.indexOf(item) === index);
  state.laundrySettings.managementEmailEnabled = !!els.laundryManagementEmailEnabled?.checked;
  state.laundrySettings.managementEmailTime = normalizeTimeInput(els.laundryManagementEmailTime?.value);
  state.laundrySettings.itemTypes = (Array.from(els.laundryItemTypesBody?.querySelectorAll("tr") || [])).map((row, index) => ({
    id: clean(state.laundrySettings.itemTypes[index]?.id) || `laundry-item-${index + 1}`,
    name: clean(row.querySelector('[data-laundry-setting-field="name"]')?.value) || `item ${index + 1}`,
    weightKg: Math.max(0, Number(normalizeNumber(row.querySelector('[data-laundry-setting-field="weightKg"]')?.value) || 0)),
  }));
  state.laundrySettings = normalizeLaundrySettingsClient(state.laundrySettings);
}

function addLaundrySettingItemType() {
  onLaundrySettingsInput();
  state.laundrySettings.itemTypes.push({
    id: `laundry-item-${Date.now()}`,
    name: `item ${state.laundrySettings.itemTypes.length + 1}`,
    weightKg: 0,
  });
  renderLaundrySettings();
}

function onLaundrySettingsAction(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) return;
  if (clean(button.dataset.action) === "remove-laundry-type") {
    onLaundrySettingsInput();
    const index = Number(button.dataset.index);
    if (Number.isFinite(index) && index >= 0) state.laundrySettings.itemTypes.splice(index, 1);
    if (!state.laundrySettings.itemTypes.length) state.laundrySettings.itemTypes = clone(DEFAULT_LAUNDRY_SETTINGS.itemTypes);
    renderLaundrySettings();
  }
}

async function saveLaundrySettings() {
  onLaundrySettingsInput();
  if (!state.laundrySettings.itemTypes.length) {
    setLaundrySettingsStatus("At least one laundry item type is required.");
    return;
  }
  try {
    const result = await api("/api/laundry-settings", { method: "PUT", body: { settings: state.laundrySettings } });
    state.laundrySettings = normalizeLaundrySettingsClient(result?.settings);
    state.laundrySettingsLoaded = true;
    state.laundryRecords = state.laundryRecords.map((row) => normalizeLaundryRecordClient(row, state.laundrySettings));
    renderLaundrySettings();
    renderLaundry();
    setLaundrySettingsStatus("Laundry configuration saved.");
    showToast("Laundry configuration saved.", "success");
  } catch (e) {
    setLaundrySettingsStatus(`Save failed: ${e.message}`);
    showToast(`Laundry configuration save failed: ${e.message}`, "error");
  }
}

function buildLaundryPayload(draft) {
  const receivedWeight = normalizeNumber(draft.receivedWeightKg);
  return {
    property: draft.property,
    date: draft.date,
    receivedDate: draft.receivedDate || laundryReceiveDate(draft.date),
    sentItems: sanitizeLaundryCountsClient(draft.sentItems),
    receivedItems: sanitizeLaundryCountsClient(draft.receivedItems),
    receivedWeightKg: receivedWeight == null ? 0 : Number(receivedWeight || 0),
    notes: clean(draft.notes),
  };
}

async function saveLaundryRecord() {
  const draft = state.laundryDraft;
  if (!draft) return;
  const validationError = validateLaundryDraftClient(draft, { isCreate: !state.laundrySelectedId });
  if (validationError) {
    setLaundryStatus(validationError);
    return;
  }
  try {
    setLaundryStatus("Saving...");
    if (state.laundrySelectedId) {
      await api(`/api/laundry?id=${encodeURIComponent(state.laundrySelectedId)}`, {
        method: "PUT",
        body: buildLaundryPayload(draft),
      });
    } else {
      await api("/api/laundry", {
        method: "POST",
        body: buildLaundryPayload(draft),
      });
    }
    await loadLaundryRecords({ silent: true });
    state.laundryLoaded = true;
    renderLaundry();
    setLaundryDbStatus("Laundry records saved.");
    closeLaundryModal();
    showToast("Laundry record saved.", "success");
  } catch (e) {
    setLaundryStatus(`Save failed: ${e.message}`);
    showToast(`Laundry record save failed: ${e.message}`, "error");
  }
}

async function onLaundryRowClick(event) {
  const row = event.target.closest("[data-laundry-id]");
  if (!row) return;
  await loadLaundryRecords({ silent: true });
  const record = state.laundryRecords.find((item) => item.id === clean(row.dataset.laundryId));
  if (!record) return;
  state.laundrySelectedId = record.id;
  state.laundryDraft = clone(record);
  openLaundryModal(record.id);
}

function render() {
  if (state.currentView === "guests") {
    renderGuests();
    return;
  }
  if (state.currentView === "lost-found") {
    renderLostFound();
    return;
  }
  if (state.currentView === "reviews") {
    renderReviews();
    return;
  }
  if (state.currentView === "groups") {
    renderGroups();
    return;
  }
  if (state.currentView === "services") {
    renderServices();
    return;
  }
  if (state.currentView === "cash") {
    renderCash();
    return;
  }
  if (state.currentView === "shopping") {
    renderShopping();
    return;
  }
  if (state.currentView === "hours") {
    renderHours();
    return;
  }
  if (state.currentView === "bakery") {
    renderBakery();
    return;
  }
  if (state.currentView === "laundry") {
    renderLaundry();
    return;
  }
  if (!canApp("communications")) {
    els.count.textContent = "0 records";
    els.rows.innerHTML = '<tr><td colspan="7" class="empty">Your profile has no access to Communications.</td></tr>';
    if (els.communicationsMobileCards) {
      els.communicationsMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Communications.</div>';
    }
    return;
  }
  const rows = getFilteredEntries();
  els.count.textContent = `${rows.length} record${rows.length === 1 ? "" : "s"}`;
  els.rows.innerHTML = "";
  if (els.communicationsMobileCards) els.communicationsMobileCards.innerHTML = "";
  els.rows.appendChild(buildInlineRow());
  renderCommunicationsMobileCards(rows);
  if (rows.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="7" class="empty">No communications found.</td>`;
    els.rows.appendChild(tr);
  }
  if (rows.length && isCommunicationsGroupingEnabled()) {
    splitCommunicationGroups(rows).forEach((group) => {
      els.rows.appendChild(buildCommunicationGroupRow(communicationGroupLabel(group)));
      group.rows.forEach((entry) => {
        els.rows.appendChild(state.editingId === entry.id ? buildEditableRow(entry) : buildReadOnlyRow(entry));
      });
    });
  } else {
    rows.forEach((entry) => els.rows.appendChild(state.editingId === entry.id ? buildEditableRow(entry) : buildReadOnlyRow(entry)));
  }
  updateSortIndicators();
  syncStickyRows();
}

function renderLostFound() {
  if (!canApp("lost-found")) {
    els.lostFoundCount.textContent = "0 records";
    els.lostFoundRows.innerHTML = '<tr><td colspan="7" class="empty">Your profile has no access to Lost&Found.</td></tr>';
    if (els.lostFoundMobileCards) {
      els.lostFoundMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Lost&Found.</div>';
    }
    return;
  }
  const rows = getFilteredLostFoundRecords();
  els.lostFoundCount.textContent = `${rows.length} record${rows.length === 1 ? "" : "s"}`;
  els.lostFoundRows.innerHTML = "";
  if (els.lostFoundMobileCards) els.lostFoundMobileCards.innerHTML = "";
  els.lostFoundRows.appendChild(buildLostFoundInlineRow());
  renderLostFoundMobileCards(rows);
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td colspan="7" class="empty">No Lost&Found records found.</td>';
    els.lostFoundRows.appendChild(tr);
    return;
  }
  rows.forEach((record) => {
    els.lostFoundRows.appendChild(
      state.lostFoundEditingId === record.id ? buildLostFoundEditableRow(record) : buildLostFoundReadOnlyRow(record)
    );
  });
}

function renderLostFoundMobileCards(rows) {
  if (!els.lostFoundMobileCards) return;
  const list = els.lostFoundMobileCards;
  list.innerHTML = "";
  list.appendChild(buildLostFoundInlineCard());
  if (!rows.length) {
    list.innerHTML += '<div class="services-mobile-empty">No Lost&Found records found.</div>';
    return;
  }
  rows.forEach((record) => {
    list.appendChild(state.lostFoundEditingId === record.id ? buildLostFoundEditableCardMobile(record) : buildLostFoundReadOnlyCardMobile(record));
  });
}

function buildLostFoundInlineRow() {
  const draft = state.lostFoundDraft || emptyLostFoundDraft();
  const now = new Date();
  const tr = document.createElement("tr");
  tr.className = "inline-editor sticky-new-row";
  tr.innerHTML = `<td class="lost-found-timestamp"><span class="auto-stamp">#${escape(nextLostFoundDisplayNumber())}</span><small>${formatDate(now)} ${formatTime(now)}</small></td>
    <td><input data-field="whoFound" data-scope="new" value="${escape(draft.whoFound)}" placeholder="Who Found" />
    <input data-field="whoRecorded" data-scope="new" value="${escape(draft.whoRecorded)}" placeholder="Who Register" /></td>
    <td><input data-field="location" data-scope="new" value="${escape(draft.location)}" placeholder="Where" />
    <select data-field="stored" data-scope="new">${LOST_FOUND_STORED_OPTIONS.map((item) => option(item, draft.stored)).join("")}</select></td>
    <td><textarea data-field="objectDescription" data-scope="new" rows="2">${escape(draft.objectDescription)}</textarea></td>
    <td><textarea data-field="notes" data-scope="new" rows="2">${escape(draft.notes)}</textarea></td>
    <td class="lost-found-status-cell"><label class="status-toggle"><input type="checkbox" data-field="status" data-scope="new" ${isClosedStatus(draft.status) ? "checked" : ""} /><span>Closed</span></label></td>
    <td class="row-actions lost-found-actions-compact"><button type="button" data-lost-found-action="save-inline">Add</button></td>`;
  tr.style.backgroundColor = "#ffffff";
  return tr;
}

function buildLostFoundReadOnlyRow(record) {
  const tr = document.createElement("tr");
  const closedStamp = isClosedStatus(record.status) && record.closedAt
    ? `<div class="status-closed-at">${escape(formatDateTimeShort(record.closedAt))}</div>`
    : "";
  tr.innerHTML = `<td class="lost-found-timestamp"><span class="lost-found-meta-strong">#${escape(record.number)}</span><small>${escape(formatDateTimeShort(record.createdAt) || "-")}</small></td>
    <td><span class="lost-found-meta-strong">Found: ${escape(record.whoFound || "-")}</span><span class="lost-found-meta-sub">Record: ${escape(record.whoRecorded || "-")}</span></td>
    <td><span class="lost-found-meta-strong">${escape(record.location || "-")}</span><span class="lost-found-meta-sub">Stored: ${escape(record.stored)}</span></td>
    <td class="lost-found-preview-cell" title="${escape(record.objectDescription)}"><span class="lost-found-notes-preview">${escape(record.objectDescription || "-")}</span></td>
    <td class="lost-found-notes-cell" title="${escape(record.notes)}"><span class="lost-found-notes-preview">${escape(record.notes || "-")}</span></td>
    <td class="lost-found-status-cell"><label class="status-toggle"><input type="checkbox" data-lost-found-action="toggle-status" data-id="${escape(record.id)}" ${isClosedStatus(record.status) ? "checked" : ""} /><span>${escape(record.status)}</span></label>${closedStamp}</td>
    <td class="row-actions lost-found-actions-compact"><button type="button" data-lost-found-action="edit" data-id="${escape(record.id)}" class="ghost">Edit</button></td>`;
  tr.style.backgroundColor = lostFoundRowBackground(record.status);
  return tr;
}

function buildLostFoundEditableRow(record) {
  const draft = state.lostFoundEditDraft || emptyLostFoundDraft();
  const tr = document.createElement("tr");
  tr.className = "inline-editor";
  tr.innerHTML = `<td class="lost-found-timestamp"><span class="lost-found-meta-strong">#${escape(record.number)}</span><small>${escape(formatDateTimeShort(record.createdAt) || "-")}</small></td>
    <td><input data-field="whoFound" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.whoFound)}" placeholder="Found" />
    <input data-field="whoRecorded" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.whoRecorded)}" placeholder="Record" /></td>
    <td><input data-field="location" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.location)}" placeholder="Where" />
    <select data-field="stored" data-scope="edit" data-id="${escape(record.id)}">${LOST_FOUND_STORED_OPTIONS.map((item) => option(item, draft.stored)).join("")}</select></td>
    <td><textarea data-field="objectDescription" data-scope="edit" data-id="${escape(record.id)}" rows="2">${escape(draft.objectDescription)}</textarea></td>
    <td><textarea data-field="notes" data-scope="edit" data-id="${escape(record.id)}" rows="2">${escape(draft.notes)}</textarea></td>
    <td class="lost-found-status-cell"><label class="status-toggle"><input type="checkbox" data-field="status" data-scope="edit" data-id="${escape(record.id)}" ${isClosedStatus(draft.status) ? "checked" : ""} /><span>Closed</span></label></td>
    <td class="row-actions lost-found-actions-compact"><button type="button" data-lost-found-action="save-edit" data-id="${escape(record.id)}">Save</button>
    <button type="button" data-lost-found-action="cancel-edit" data-id="${escape(record.id)}" class="ghost">Cancel</button></td>`;
  tr.style.backgroundColor = lostFoundRowBackground(draft.status);
  return tr;
}

function buildLostFoundInlineCard() {
  const draft = state.lostFoundDraft || emptyLostFoundDraft();
  const now = new Date();
  const card = document.createElement("article");
  card.className = "lost-found-mobile-card";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">#${escape(nextLostFoundDisplayNumber())}</div>
        <div class="communication-mobile-meta">${escape(formatDate(now))} ${escape(formatTime(now))}</div>
      </div>
      <label class="status-toggle"><input type="checkbox" data-field="status" data-scope="new" ${isClosedStatus(draft.status) ? "checked" : ""} /><span>Closed</span></label>
    </div>
    <div class="communication-mobile-grid">
      <label class="communication-mobile-field">
        <small>Who Found</small>
        <input data-field="whoFound" data-scope="new" value="${escape(draft.whoFound)}" placeholder="Who Found" />
      </label>
      <label class="communication-mobile-field">
        <small>Who Register</small>
        <input data-field="whoRecorded" data-scope="new" value="${escape(draft.whoRecorded)}" placeholder="Who Register" />
      </label>
      <label class="communication-mobile-field">
        <small>Where</small>
        <input data-field="location" data-scope="new" value="${escape(draft.location)}" placeholder="Where" />
      </label>
      <label class="communication-mobile-field">
        <small>Stored</small>
        <select data-field="stored" data-scope="new">${LOST_FOUND_STORED_OPTIONS.map((item) => option(item, draft.stored)).join("")}</select>
      </label>
      <label class="communication-mobile-field communication-mobile-field-full">
        <small>Object Description</small>
        <textarea data-field="objectDescription" data-scope="new" rows="3">${escape(draft.objectDescription)}</textarea>
      </label>
      <label class="communication-mobile-field communication-mobile-field-full">
        <small>Notes</small>
        <textarea data-field="notes" data-scope="new" rows="3">${escape(draft.notes)}</textarea>
      </label>
    </div>
    <div class="communication-mobile-footer">
      <div class="row-actions"><button type="button" data-lost-found-action="save-inline">Add</button></div>
    </div>`;
  card.style.backgroundColor = "#ffffff";
  return card;
}

function buildLostFoundReadOnlyCardMobile(record) {
  const closedStamp = isClosedStatus(record.status) && record.closedAt
    ? `<div class="status-closed-at">${escape(formatDateTimeShort(record.closedAt))}</div>`
    : "";
  const card = document.createElement("article");
  card.className = "lost-found-mobile-card";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">#${escape(record.number)}</div>
        <div class="communication-mobile-meta">${escape(formatDateTimeShort(record.createdAt) || "-")}</div>
      </div>
      <div class="communication-mobile-status">
        <label class="status-toggle"><input type="checkbox" data-lost-found-action="toggle-status" data-id="${escape(record.id)}" ${isClosedStatus(record.status) ? "checked" : ""} /><span>${escape(record.status)}</span></label>
        ${closedStamp}
      </div>
    </div>
    <div class="communication-mobile-grid">
      <div class="communication-mobile-field">
        <small>Who Found</small>
        <div class="communication-mobile-message">${escape(record.whoFound || "-")}</div>
      </div>
      <div class="communication-mobile-field">
        <small>Who Register</small>
        <div class="communication-mobile-message">${escape(record.whoRecorded || "-")}</div>
      </div>
      <div class="communication-mobile-field">
        <small>Where</small>
        <div class="communication-mobile-message">${escape(record.location || "-")}</div>
      </div>
      <div class="communication-mobile-field">
        <small>Stored</small>
        <div class="communication-mobile-message">${escape(record.stored || "-")}</div>
      </div>
      <div class="communication-mobile-field communication-mobile-field-full">
        <small>Object Description</small>
        <div class="communication-mobile-message">${escape(record.objectDescription || "-")}</div>
      </div>
      <div class="communication-mobile-field communication-mobile-field-full">
        <small>Notes</small>
        <div class="communication-mobile-message">${escape(record.notes || "-")}</div>
      </div>
    </div>
    <div class="communication-mobile-footer">
      <div class="row-actions"><button type="button" data-lost-found-action="edit" data-id="${escape(record.id)}" class="ghost">Edit</button></div>
    </div>`;
  card.style.backgroundColor = lostFoundRowBackground(record.status);
  return card;
}

function buildLostFoundEditableCardMobile(record) {
  const draft = state.lostFoundEditDraft || emptyLostFoundDraft();
  const card = document.createElement("article");
  card.className = "lost-found-mobile-card";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">#${escape(record.number)}</div>
        <div class="communication-mobile-meta">${escape(formatDateTimeShort(record.createdAt) || "-")}</div>
      </div>
      <label class="status-toggle"><input type="checkbox" data-field="status" data-scope="edit" data-id="${escape(record.id)}" ${isClosedStatus(draft.status) ? "checked" : ""} /><span>Closed</span></label>
    </div>
    <div class="communication-mobile-grid">
      <label class="communication-mobile-field">
        <small>Who Found</small>
        <input data-field="whoFound" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.whoFound)}" placeholder="Who Found" />
      </label>
      <label class="communication-mobile-field">
        <small>Who Register</small>
        <input data-field="whoRecorded" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.whoRecorded)}" placeholder="Who Register" />
      </label>
      <label class="communication-mobile-field">
        <small>Where</small>
        <input data-field="location" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.location)}" placeholder="Where" />
      </label>
      <label class="communication-mobile-field">
        <small>Stored</small>
        <select data-field="stored" data-scope="edit" data-id="${escape(record.id)}">${LOST_FOUND_STORED_OPTIONS.map((item) => option(item, draft.stored)).join("")}</select>
      </label>
      <label class="communication-mobile-field communication-mobile-field-full">
        <small>Object Description</small>
        <textarea data-field="objectDescription" data-scope="edit" data-id="${escape(record.id)}" rows="3">${escape(draft.objectDescription)}</textarea>
      </label>
      <label class="communication-mobile-field communication-mobile-field-full">
        <small>Notes</small>
        <textarea data-field="notes" data-scope="edit" data-id="${escape(record.id)}" rows="3">${escape(draft.notes)}</textarea>
      </label>
    </div>
    <div class="communication-mobile-footer">
      <div class="row-actions">
        <button type="button" data-lost-found-action="save-edit" data-id="${escape(record.id)}">Save</button>
        <button type="button" data-lost-found-action="cancel-edit" data-id="${escape(record.id)}" class="ghost">Cancel</button>
      </div>
    </div>`;
  card.style.backgroundColor = lostFoundRowBackground(draft.status);
  return card;
}

function buildInlineRow() {
  const tr = document.createElement("tr");
  tr.className = "inline-editor sticky-new-row";
  const now = new Date();
  tr.innerHTML = `<td><span class="auto-stamp">New row</span><div>${formatDate(now)}</div></td>
    <td>${formatTime(now)}</td>
    <td><input data-field="person" data-scope="new" value="${escape(state.newDraft.person)}" /></td>
    <td><select data-field="category" data-scope="new">${getCategories().map((c) => option(c.name, state.newDraft.category)).join("")}</select></td>
    <td><textarea data-field="message" data-scope="new" rows="2">${escape(state.newDraft.message)}</textarea></td>
    <td><label class="status-toggle"><input type="checkbox" data-field="status" data-scope="new" ${isClosedStatus(state.newDraft.status) ? "checked" : ""} /><span>Closed</span></label></td>
    <td class="row-actions"><button type="button" data-action="save-inline">Add</button></td>`;
  tr.style.backgroundColor = "#ffffff";
  return tr;
}

function buildReadOnlyRow(entry) {
  const tr = document.createElement("tr");
  const closedStamp = isClosedStatus(entry.status) ? `<div class="status-closed-at">${escape(formatDateTimeShort(entry.updatedAt || entry.createdAt))}</div>` : "";
  tr.innerHTML = `<td>${escape(entry.date)}</td><td>${escape(entry.time)}</td><td>${escape(entry.person)}</td>
    <td><span class="chip" style="${chipStyle(getCategory(entry.category).color)}">${escape(entry.category)}</span></td>
    <td class="message">${escape(entry.message)}</td>
    <td><label class="status-toggle"><input type="checkbox" data-action="toggle-status" data-id="${entry.id}" ${isClosedStatus(entry.status) ? "checked" : ""} /><span>${escape(entry.status)}</span></label>${closedStamp}</td>
    <td class="row-actions"><button data-action="edit" data-id="${entry.id}">Edit</button>
    <button data-action="delete" data-id="${entry.id}" class="danger">Delete</button></td>`;
  tr.style.backgroundColor = rowBackgroundColor(entry.status, entry.category);
  return tr;
}

function buildEditableRow(entry) {
  const d = state.editDraft;
  const tr = document.createElement("tr");
  tr.className = "inline-editor";
  tr.innerHTML = `<td>${escape(entry.date)}</td><td>${escape(entry.time)}</td>
    <td><input data-field="person" data-scope="edit" data-id="${entry.id}" value="${escape(d.person)}" /></td>
    <td><select data-field="category" data-scope="edit" data-id="${entry.id}">${getCategories().map((c) => option(c.name, d.category)).join("")}</select></td>
    <td><textarea data-field="message" data-scope="edit" data-id="${entry.id}" rows="2">${escape(d.message)}</textarea></td>
    <td><label class="status-toggle"><input type="checkbox" data-field="status" data-scope="edit" data-id="${entry.id}" ${isClosedStatus(d.status) ? "checked" : ""} /><span>Closed</span></label></td>
    <td class="row-actions"><button data-action="save-edit" data-id="${entry.id}">Save</button>
    <button data-action="cancel-edit" data-id="${entry.id}" class="ghost">Cancel</button></td>`;
  tr.style.backgroundColor = rowBackgroundColor(d.status, d.category);
  return tr;
}

function buildCommunicationGroupRow(label) {
  const tr = document.createElement("tr");
  tr.className = "communications-group-row";
  tr.innerHTML = `<td colspan="7"><span class="communications-group-pill">${escape(label)}</span></td>`;
  return tr;
}

function renderCommunicationsMobileCards(rows) {
  if (!els.communicationsMobileCards) return;
  const list = els.communicationsMobileCards;
  list.innerHTML = "";
  list.appendChild(buildCommunicationInlineCard());
  if (!rows.length) {
    list.innerHTML += '<div class="services-mobile-empty">No communications found.</div>';
    return;
  }
  if (isCommunicationsGroupingEnabled()) {
    splitCommunicationGroups(rows).forEach((group) => {
      list.appendChild(buildCommunicationGroupCard(communicationGroupLabel(group)));
      group.rows.forEach((entry) => {
        list.appendChild(state.editingId === entry.id ? buildCommunicationEditableCard(entry) : buildCommunicationReadOnlyCard(entry));
      });
    });
    return;
  }
  rows.forEach((entry) => {
    list.appendChild(state.editingId === entry.id ? buildCommunicationEditableCard(entry) : buildCommunicationReadOnlyCard(entry));
  });
}

function buildCommunicationGroupCard(label) {
  const card = document.createElement("div");
  card.className = "communications-mobile-group";
  card.textContent = label;
  return card;
}

function buildCommunicationInlineCard() {
  const now = new Date();
  const card = document.createElement("article");
  card.className = "communication-mobile-card communication-mobile-inline";
  card.innerHTML = `<div class="communication-mobile-header">
      <div class="service-mobile-request">New row</div>
      <div class="communication-mobile-meta">${escape(formatDate(now))} ${escape(formatTime(now))}</div>
    </div>
    <div class="communication-mobile-grid">
      <label class="communication-mobile-field">
        <small>Person</small>
        <input data-field="person" data-scope="new" value="${escape(state.newDraft.person)}" />
      </label>
      <label class="communication-mobile-field">
        <small>Category</small>
        <select data-field="category" data-scope="new">${getCategories().map((c) => option(c.name, state.newDraft.category)).join("")}</select>
      </label>
      <label class="communication-mobile-field communication-mobile-field-full">
        <small>What happened</small>
        <textarea data-field="message" data-scope="new" rows="3">${escape(state.newDraft.message)}</textarea>
      </label>
    </div>
    <div class="communication-mobile-footer">
      <label class="status-toggle"><input type="checkbox" data-field="status" data-scope="new" ${isClosedStatus(state.newDraft.status) ? "checked" : ""} /><span>Closed</span></label>
      <div class="row-actions"><button type="button" data-action="save-inline">Add</button></div>
    </div>`;
  card.style.backgroundColor = "#ffffff";
  return card;
}

function buildCommunicationReadOnlyCard(entry) {
  const card = document.createElement("article");
  const closedStamp = isClosedStatus(entry.status) ? `<div class="status-closed-at">${escape(formatDateTimeShort(entry.updatedAt || entry.createdAt))}</div>` : "";
  card.className = "communication-mobile-card";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(entry.person)}</div>
        <div class="communication-mobile-meta">${escape(entry.date)} ${escape(entry.time)}</div>
      </div>
      <span class="chip" style="${chipStyle(getCategory(entry.category).color)}">${escape(entry.category)}</span>
    </div>
    <div class="communication-mobile-message">${escape(entry.message)}</div>
    <div class="communication-mobile-footer">
      <div class="communication-mobile-status">
        <label class="status-toggle"><input type="checkbox" data-action="toggle-status" data-id="${entry.id}" ${isClosedStatus(entry.status) ? "checked" : ""} /><span>${escape(entry.status)}</span></label>
        ${closedStamp}
      </div>
      <div class="row-actions">
        <button data-action="edit" data-id="${entry.id}">Edit</button>
        <button data-action="delete" data-id="${entry.id}" class="danger">Delete</button>
      </div>
    </div>`;
  card.style.backgroundColor = rowBackgroundColor(entry.status, entry.category);
  return card;
}

function buildCommunicationEditableCard(entry) {
  const d = state.editDraft;
  const card = document.createElement("article");
  card.className = "communication-mobile-card communication-mobile-editing";
  card.innerHTML = `<div class="communication-mobile-header">
      <div>
        <div class="service-mobile-request">${escape(entry.date)} ${escape(entry.time)}</div>
        <div class="communication-mobile-meta">Editing communication</div>
      </div>
      <label class="status-toggle"><input type="checkbox" data-field="status" data-scope="edit" data-id="${entry.id}" ${isClosedStatus(d.status) ? "checked" : ""} /><span>Closed</span></label>
    </div>
    <div class="communication-mobile-grid">
      <label class="communication-mobile-field">
        <small>Person</small>
        <input data-field="person" data-scope="edit" data-id="${entry.id}" value="${escape(d.person)}" />
      </label>
      <label class="communication-mobile-field">
        <small>Category</small>
        <select data-field="category" data-scope="edit" data-id="${entry.id}">${getCategories().map((c) => option(c.name, d.category)).join("")}</select>
      </label>
      <label class="communication-mobile-field communication-mobile-field-full">
        <small>What happened</small>
        <textarea data-field="message" data-scope="edit" data-id="${entry.id}" rows="3">${escape(d.message)}</textarea>
      </label>
    </div>
    <div class="communication-mobile-footer">
      <div class="row-actions">
        <button data-action="save-edit" data-id="${entry.id}">Save</button>
        <button data-action="cancel-edit" data-id="${entry.id}" class="ghost">Cancel</button>
      </div>
    </div>`;
  card.style.backgroundColor = rowBackgroundColor(d.status, d.category);
  return card;
}

async function importFromExcel(event) {
  const file = event.target.files?.[0];
  if (!file || !window.XLSX) return;
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const workbook = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
      if (!workbook.Sheets[SHEET_NAME]) return showToast(`Sheet "${SHEET_NAME}" not found.`, "error");
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[SHEET_NAME], { header: 1, raw: false, defval: "" });
      const payload = parseSheetRows(rows).map(({ id, ...x }) => x);
      if (payload.length === 0) return showToast("No valid rows found.", "error");
      await api("/api/communications", { method: "POST", body: payload });
      await loadEntries();
      render();
      showToast(`Imported ${payload.length} communications.`, "success");
    } catch (e2) {
      showToast(`Import failed: ${e2.message}`, "error");
    } finally {
      els.excelInput.value = "";
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseSheetRows(rows) {
  const h = rows.findIndex((r) => r.map((x) => clean(x).toLowerCase()).includes("pessoa"));
  if (h === -1) return [];
  const header = rows[h].map((c) => clean(c).toLowerCase());
  const col = {
    date: header.indexOf("data"),
    time: header.indexOf("hora"),
    person: header.indexOf("pessoa"),
    message: header.findIndex((x) => x.includes("o que aconteceu")),
    status: header.indexOf("status"),
    category: Math.max(header.indexOf("category"), header.indexOf("categoria")),
  };
  return rows.slice(h + 1).map((row) => ({
    id: crypto.randomUUID(),
    date: normalizeDate(clean(row[col.date])),
    time: normalizeTime(clean(row[col.time])),
    person: clean(row[col.person]),
    status: normalizeStatusUi(clean(row[col.status])),
    category: normalizeCategory(clean(row[col.category])),
    message: clean(row[col.message]),
  })).filter((x) => x.person || x.message);
}

function exportToCsv() {
  const header = ["Data", "Hora", "Pessoa", "Status", "Category", "O que aconteceu?"];
  const lines = getFilteredEntries().map((e) => [e.date, e.time, e.person, e.status, e.category, e.message]);
  downloadCsv("communications_log.csv", [header, ...lines]);
  showToast(`Exported ${lines.length} communications.`, "success");
}

function exportReviewsToCsv() {
  const header = [
    "Review date",
    "Property",
    "Source",
    "Reviewer",
    "Reviewer country",
    "Language",
    "Title",
    "Review text",
    "Positive review",
    "Negative review",
    "Rating normalized 100",
    "Rating raw",
    "Rating scale",
    "Staff",
    "Cleanliness",
    "Location",
    "Facilities",
    "Comfort",
    "Value for money",
    "Rooms",
    "Service",
    "Sleep quality",
    "Booking / Reservation",
    "Source reference",
    "Property reply",
    "Host reply date",
  ];
  const rows = getFilteredReviews();
  const lines = rows.map((row) => {
    const subscores = row.subscores || {};
    return [
      clean(row.review_date),
      clean(row.properties?.name || reviewPropertyName(row.property_id)),
      reviewSourceLabel(row.source),
      clean(row.reviewer_name),
      clean(row.reviewer_country),
      clean(row.language),
      clean(row.title),
      buildReviewBodyPreview(row),
      clean(row.positive_review_text),
      clean(row.negative_review_text),
      row.rating_normalized_100 ?? "",
      row.rating_raw ?? "",
      row.rating_scale ?? "",
      subscores.staff ?? "",
      subscores.cleanliness ?? "",
      subscores.location ?? "",
      subscores.facilities ?? "",
      subscores.comfort ?? "",
      subscores.value_for_money ?? "",
      subscores.rooms ?? "",
      subscores.service ?? "",
      subscores.sleep_quality ?? "",
      clean(row.source_reservation_id),
      clean(row.source_review_id),
      clean(row.host_reply_text),
      clean(row.host_reply_date),
    ];
  });
  const date = formatDate(new Date());
  downloadCsv(`reviews_export_${date}.csv`, [header, ...lines]);
  showToast(`Exported ${lines.length} reviews.`, "success");
}

function downloadCsv(fileName, rows) {
  const csv = rows.map((row) => row.map(csvCell).join(",")).join("\n");
  downloadBlob(fileName, csv, "text/csv;charset=utf-8;");
}

function downloadBlob(fileName, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function csvCell(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function sanitizeSettings(settings) {
  const output = clone(DEFAULT_SETTINGS);
  const generalInput = settings?.general || {};
  output.general.emailConfig = normalizeBakeryEmailConfigClient(generalInput.emailConfig || generalInput.email_config || generalInput.bakeryEmailConfig || generalInput.bakery_email_config);
  const input = settings?.communications || {};
  const categories = Array.isArray(input.categories) ? input.categories : [];
  const seen = new Set();
  const hasTaskCategory = categories.some((c) => clean(c?.name).toLowerCase() === "task");
  output.communications.categories = categories
    .map((c) => {
      const rawName = clean(c?.name);
      const lowered = rawName.toLowerCase();
      const name = lowered === "very important" && !hasTaskCategory ? "Task" : rawName;
      return { name, color: normalizeHex(c?.color), autoCloseDays: normalizeAutoCloseDays(c?.autoCloseDays) };
    })
    .filter((c) => c.name)
    .filter((c) => (seen.has(c.name.toLowerCase()) ? false : (seen.add(c.name.toLowerCase()), true)));
  if (output.communications.categories.length === 0) output.communications.categories = clone(DEFAULT_SETTINGS).communications.categories;
  output.communications.emailAutomation.enabled = !!input.emailAutomation?.enabled;
  output.communications.emailAutomation.frequency = normalizeFrequency(input.emailAutomation?.frequency);
  output.communications.emailAutomation.timeOfDay = normalizeTimeInput(input.emailAutomation?.timeOfDay);
  output.communications.emailAutomation.recipients = parseEmailList(input.emailAutomation?.recipients);
  output.communications.emailAutomation.frequency2 = normalizeFrequency(input.emailAutomation?.frequency2);
  output.communications.emailAutomation.timeOfDay2 = normalizeTimeInput(input.emailAutomation?.timeOfDay2);
  output.communications.emailAutomation.recipients2 = parseEmailList(input.emailAutomation?.recipients2);
  return output;
}

function normalizeDraftsToSettings() {
  const list = getCategories();
  const fix = (val) => list.find((x) => x.name.toLowerCase() === clean(val).toLowerCase())?.name || list[0].name;
  state.entries = state.entries.map((e) => ({ ...e, category: fix(e.category) }));
  state.newDraft.category = fix(state.newDraft.category);
  if (state.editDraft) state.editDraft.category = fix(state.editDraft.category);
}

function getCategories() {
  return state.settings.communications.categories;
}

function normalizeCategory(value) {
  return getCategory(value).name;
}

function getCategory(name) {
  const raw = clean(name).toLowerCase();
  return getCategories().find((x) => x.name.toLowerCase() === raw) || getCategories()[0];
}

function chipStyle(color) {
  const bg = normalizeHex(color);
  return `background:${bg};border-color:${bg};color:${contrastText(bg)};`;
}

function contrastText(hex) {
  const x = normalizeHex(hex).slice(1);
  const r = parseInt(x.slice(0, 2), 16);
  const g = parseInt(x.slice(2, 4), 16);
  const b = parseInt(x.slice(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.6 ? "#1d1714" : "#ffffff";
}

function normalizeHex(value) {
  const raw = clean(value);
  if (/^#[0-9a-fA-F]{6}$/.test(raw)) return raw.toLowerCase();
  if (/^#[0-9a-fA-F]{3}$/.test(raw)) return `#${raw[1]}${raw[1]}${raw[2]}${raw[2]}${raw[3]}${raw[3]}`.toLowerCase();
  return "#d8d8d8";
}

function normalizeAutoCloseDays(value) {
  const raw = clean(value);
  if (!raw) return null;
  const num = Number(raw);
  if (!Number.isFinite(num) || num <= 0) return null;
  return Math.floor(num);
}

function normalizeFrequency(value) {
  return value === "every1h" || value === "every4h" || value === "every8h" ? value : "everyday";
}

function emailFrequencyStep(value) {
  if (value === "every1h") return 1;
  if (value === "every4h") return 4;
  return 8;
}

function communicationEmailSchedules(email) {
  const source = email || {};
  return [
    {
      key: "schedule1",
      label: "Schedule 1",
      frequency: normalizeFrequency(source.frequency),
      timeOfDay: normalizeTimeInput(source.timeOfDay),
      recipients: parseEmailList(source.recipients),
    },
    {
      key: "schedule2",
      label: "Schedule 2",
      frequency: normalizeFrequency(source.frequency2),
      timeOfDay: normalizeTimeInput(source.timeOfDay2),
      recipients: parseEmailList(source.recipients2),
    },
  ];
}

function emailScheduleSummary(schedule) {
  const recipients = schedule.recipients;
  if (!recipients.length) return "";
  if (schedule.frequency === "everyday") {
    return `${schedule.label}: ${recipients.length} recipient${recipients.length === 1 ? "" : "s"} every day at ${schedule.timeOfDay}.`;
  }
  const step = emailFrequencyStep(schedule.frequency);
  const [h, m] = schedule.timeOfDay.split(":");
  const start = Number(h) || 0;
  const items = [];
  for (let hour = start; hour < 24; hour += step) items.push(`${String(hour).padStart(2, "0")}:${m || "00"}`);
  return `${schedule.label}: ${recipients.length} recipient${recipients.length === 1 ? "" : "s"} at ${items.join(", ")} every day.`;
}

function emailScheduleSummaries(email) {
  return communicationEmailSchedules(email)
    .map(emailScheduleSummary)
    .filter(Boolean);
}

function normalizeDateInput(value) {
  const raw = clean(value);
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const dmyMatch = raw.match(/^(\d{2})[\/.-](\d{2})[\/.-](\d{4})$/);
  if (dmyMatch) return `${dmyMatch[3]}-${dmyMatch[2]}-${dmyMatch[1]}`;
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return new Intl.DateTimeFormat("sv-SE", {
      timeZone: "Europe/Lisbon",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(parsed);
  }
  return raw.slice(0, 10);
}

function normalizeTimeInput(value) {
  const raw = clean(value);
  if (/^\d{2}:\d{2}$/.test(raw)) return raw;
  if (/^\d{2}:\d{2}:\d{2}$/.test(raw)) return raw.slice(0, 5);
  return "00:00";
}

function normalizeStatusUi(value) {
  const raw = clean(value).toLowerCase();
  if (raw === "closed" || raw === "resolved" || raw === "archived") return "Closed";
  return "Open";
}

function parseEmailList(value) {
  const raw = Array.isArray(value) ? value.join(",") : clean(value);
  const seen = new Set();
  return raw
    .split(/[\n,;]/)
    .map((x) => clean(x).toLowerCase())
    .filter((x) => isValidEmail(x))
    .filter((x) => (seen.has(x) ? false : (seen.add(x), true)));
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean(value));
}

function isClosedStatus(value) {
  return normalizeStatusUi(value) === "Closed";
}

function rowBackgroundColor(status, category) {
  if (isClosedStatus(status)) return hexToRgba("#2e9f42", 0.25);
  return hexToRgba(getCategory(category).color, 0.25);
}

function hexToRgba(hex, alpha) {
  const raw = normalizeHex(hex).slice(1);
  const r = parseInt(raw.slice(0, 2), 16);
  const g = parseInt(raw.slice(2, 4), 16);
  const b = parseInt(raw.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function option(label, selected) {
  const sel = clean(label) === clean(selected) ? "selected" : "";
  return `<option value="${escape(label)}" ${sel}>${escape(label)}</option>`;
}

function queuePendingDelete(entry, index) {
  clearPendingDelete();
  state.pendingDelete = {
    entry: clone(entry),
    index,
    timer: window.setTimeout(() => {
      state.pendingDelete = null;
    }, 9000),
  };
}

function clearPendingDelete() {
  if (!state.pendingDelete?.timer) return;
  clearTimeout(state.pendingDelete.timer);
  state.pendingDelete = null;
}

async function undoPendingDelete() {
  const pending = state.pendingDelete;
  if (!pending?.entry) return;
  clearPendingDelete();
  const item = pending.entry;
  try {
    await api("/api/communications", {
      method: "POST",
      body: {
        date: normalizeDate(item.date),
        time: normalizeTime(item.time),
        person: clean(item.person),
        status: normalizeStatusUi(item.status),
        category: normalizeCategory(item.category),
        message: clean(item.message),
      },
    });
    await loadEntries();
    render();
    showToast("Record restored.", "success");
  } catch (error) {
    showToast(`Undo failed: ${error.message}`, "error");
  }
}

function ensureToastHost() {
  let host = document.getElementById("toast-host");
  if (!host) {
    host = document.createElement("div");
    host.id = "toast-host";
    host.className = "toast-host";
    document.body.appendChild(host);
  }
  els.toastHost = host;
}

function showToast(message, type = "info", options = {}) {
  ensureToastHost();
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  const text = document.createElement("span");
  text.textContent = message;
  toast.appendChild(text);

  if (options.actionLabel && typeof options.action === "function") {
    const actionBtn = document.createElement("button");
    actionBtn.type = "button";
    actionBtn.className = "ghost";
    actionBtn.textContent = options.actionLabel;
    actionBtn.addEventListener("click", () => {
      options.action();
      toast.remove();
    });
    toast.appendChild(actionBtn);
  }

  els.toastHost.appendChild(toast);
  const duration = Number(options.duration) > 0 ? Number(options.duration) : 5000;
  window.setTimeout(() => toast.remove(), duration);
}

function setDbStatus(text) {
  els.dbStatus.textContent = text;
}

function setSettingsStatus(text) {
  els.settingsStatus.textContent = text;
}

function setGeneralSettingsStatus(text) {
  if (els.generalSettingsStatus) els.generalSettingsStatus.textContent = text;
}

function setAdminUsersStatus(text) {
  els.adminUsersStatus.textContent = text;
}

function setProfilesStatus(text) {
  els.profilesStatus.textContent = text;
}

function normalizeDate(value) {
  if (!value) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  const dmyMatch = String(value).match(/^(\d{2})[\/.-](\d{2})[\/.-](\d{4})$/);
  if (dmyMatch) return `${dmyMatch[3]}-${dmyMatch[2]}-${dmyMatch[1]}`;
  const portugueseDate = normalizePortugueseDate(value);
  if (portugueseDate) return portugueseDate;
  const dt = new Date(value);
  return Number.isNaN(dt.getTime()) ? value : formatDate(dt);
}

function normalizePortugueseDate(value) {
  const raw = clean(value).toLowerCase().replace(/[\u2009\u202f]/g, " ").replace(/\s+/g, " ");
  const dateSeparator = "[\\s\\u2009\\u202f]*[^\\da-z.]+[\\s\\u2009\\u202f]*";
  const crossYearMatch = raw.match(new RegExp(`(\\d{1,2})\\s+de\\s+([a-zç.]+)\\s+de\\s+\\d{4}${dateSeparator}(\\d{1,2})\\s+de\\s+([a-zç.]+)\\s+de\\s+(\\d{4})`, "i"));
  if (crossYearMatch) {
    return formatPortugueseDateParts(crossYearMatch[3], crossYearMatch[4], crossYearMatch[5]);
  }
  const crossMonthMatch = raw.match(new RegExp(`(\\d{1,2})\\s+de\\s+([a-zç.]+)${dateSeparator}(\\d{1,2})\\s+de\\s+([a-zç.]+)\\s+de\\s+(\\d{4})`, "i"));
  if (crossMonthMatch) {
    return formatPortugueseDateParts(crossMonthMatch[3], crossMonthMatch[4], crossMonthMatch[5]);
  }
  const match = raw.match(new RegExp(`(\\d{1,2})(?:${dateSeparator}(\\d{1,2}))?\\s+(?:de\\s+)?([a-zç.]+)\\s+de\\s+(\\d{4})`, "i"));
  if (!match) return "";
  const day = match[2] || match[1];
  return formatPortugueseDateParts(day, match[3], match[4]);
}

function formatPortugueseDateParts(day, monthName, year) {
  const monthKey = clean(monthName).toLowerCase().replace(".", "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const months = {
    jan: 1, janeiro: 1,
    fev: 2, fevereiro: 2,
    mar: 3, marco: 3,
    abr: 4, abril: 4,
    mai: 5, maio: 5,
    jun: 6, junho: 6,
    jul: 7, julho: 7,
    ago: 8, agosto: 8,
    set: 9, setembro: 9,
    out: 10, outubro: 10,
    nov: 11, novembro: 11,
    dez: 12, dezembro: 12,
  };
  const month = months[monthKey];
  if (!month) return "";
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function normalizeTime(value) {
  if (!value) return "";
  if (/^\d{2}:\d{2}/.test(value)) return value.slice(0, 5);
  const dt = new Date(`1970-01-01T${value}`);
  return Number.isNaN(dt.getTime()) ? value : formatTime(dt);
}

function formatDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatTime(date) {
  return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

function formatDateTimeShort(value) {
  const raw = clean(value);
  if (!raw) return "-";
  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return raw;
  return `${formatDate(d)} ${formatTime(d)}`;
}

function openReviewsSettings() {
  setView("settings");
  setSettingsSection("reviews");
  if (!state.reviewStagingRows.length) loadLatestParsedImportRun();
}

function applyDefaultReviewDateFilter() {
  if (state.reviewFilters.dateFrom || clean(els.reviewsFromDate?.value)) return;
  const today = new Date();
  const defaultFrom = new Date(today.getFullYear() - 1, 0, 1);
  state.reviewFilters.dateFrom = formatDate(defaultFrom);
  if (els.reviewsFromDate) els.reviewsFromDate.value = state.reviewFilters.dateFrom;
}

function setReviewScreen(screen) {
  state.reviewScreen = screen === "resume" || screen === "rating" ? screen : "list";
  els.reviewsScreenList.classList.toggle("active-tab", state.reviewScreen === "list");
  els.reviewsScreenResume.classList.toggle("active-tab", state.reviewScreen === "resume");
  els.reviewsScreenRating.classList.toggle("active-tab", state.reviewScreen === "rating");
  els.reviewsScreenPanelList.hidden = state.reviewScreen !== "list";
  els.reviewsScreenPanelResume.hidden = state.reviewScreen !== "resume";
  els.reviewsScreenPanelRating.hidden = state.reviewScreen !== "rating";
  if (state.reviewScreen !== "list") closeReviewDetailModal();
  renderReviewQa();
}

function setReviewSettingsScreen(screen, loadData = true) {
  state.reviewSettingsScreen = screen === "config" ? "config" : "import";
  const isImport = state.reviewSettingsScreen === "import";
  els.settingsReviewsImportTab.classList.toggle("active-tab", isImport);
  els.settingsReviewsImportTab.classList.toggle("ghost", !isImport);
  els.settingsReviewsConfigTab.classList.toggle("active-tab", !isImport);
  els.settingsReviewsConfigTab.classList.toggle("ghost", isImport);
  els.settingsReviewsImportPanel.hidden = !isImport;
  els.settingsReviewsConfigPanel.hidden = isImport;
  if (loadData && state.currentView === "settings" && state.settingsSection === "reviews") {
    if (isImport && !state.reviewImportRunsLoaded) {
      loadReviewImportRuns().then(() => {
        state.reviewImportRunsLoaded = true;
        renderReviewImportRuns();
      }).catch((e) => setReviewImportStatus(`Could not load recent imports: ${e.message}`));
    }
    if (!isImport && !state.reviewGoogleLoaded) {
      loadGoogleBusinessStatus().then(() => {
        state.reviewGoogleLoaded = true;
        renderGoogleBusinessSettings();
      });
    }
  }
}

function onReviewFilterInput() {
  state.reviewFilters.propertyId = clean(els.reviewsPropertyFilter.value);
  state.reviewFilters.source = clean(els.reviewsSourceFilter.value);
  state.reviewFilters.search = clean(els.reviewsSearch.value).toLowerCase();
  state.reviewFilters.dateFrom = clean(els.reviewsFromDate.value);
  state.reviewFilters.dateTo = clean(els.reviewsToDate.value);
  state.reviewFilters.scoreFrom = clean(els.reviewsScoreFrom.value);
  state.reviewFilters.scoreTo = clean(els.reviewsScoreTo.value);
  state.reviewListPage = 1;
  if (state.reviewScreen !== "list") state.reviewSelectedId = "";
  renderReviews();
}

async function loadReviewProperties() {
  try {
    const result = await api("/api/properties");
    state.reviewProperties = Array.isArray(result.rows) ? result.rows : [];
    renderReviewProperties();
  } catch (e) {
    state.reviewProperties = [];
    setReviewsStatus(`Failed to load properties: ${e.message}`);
  }
}

function renderReviewPropertyOptions() {
  const currentFilter = clean(state.reviewFilters.propertyId || els.reviewsPropertyFilter.value);
  const currentImport = clean(els.reviewsImportProperty.value);
  const options = ['<option value="">All properties</option>']
    .concat(state.reviewProperties.filter((row) => row.active !== false).map((row) => `<option value="${escape(row.id)}">${escape(row.name)}</option>`))
    .join("");
  els.reviewsPropertyFilter.innerHTML = options;
  els.reviewsImportProperty.innerHTML = ['<option value="">Select property</option>']
    .concat(state.reviewProperties
      .filter((row) => row.active !== false)
      .map((row) => `<option value="${escape(row.id)}">${escape(row.name)}</option>`))
    .join("");
  els.reviewsPropertyFilter.value = currentFilter;
  els.reviewsImportProperty.value = currentImport;
  renderReviewSourceOptions();
}

async function loadReviews({ useFilters = false, silent = false } = {}) {
  try {
    const rows = [];
    let offset = 0;
    const filterQuery = useFilters ? reviewApiFilterQuery() : "";
    while (true) {
      const result = await api(`/api/reviews?limit=${REVIEW_FETCH_PAGE_SIZE}&offset=${offset}${filterQuery}`);
      const pageRows = Array.isArray(result.rows) ? result.rows : [];
      if (pageRows.length === 0) break;
      rows.push(...pageRows);
      offset += pageRows.length;
    }
    state.reviews = rows;
    if (!silent) setReviewsStatus(`Loaded ${state.reviews.length} reviews.`);
  } catch (e) {
    state.reviews = [];
    setReviewsStatus(`Failed to load reviews: ${e.message}`);
    showToast(`Failed to load reviews: ${e.message}`, "error");
  }
}

function reviewApiFilterQuery() {
  const params = new URLSearchParams();
  const filters = state.reviewFilters;
  if (clean(filters.propertyId)) params.set("propertyId", clean(filters.propertyId));
  if (clean(filters.source)) params.set("source", clean(filters.source));
  if (clean(filters.search)) params.set("search", clean(filters.search));
  if (clean(filters.dateFrom)) params.set("dateFrom", clean(filters.dateFrom));
  if (clean(filters.dateTo)) params.set("dateTo", clean(filters.dateTo));
  const query = params.toString();
  return query ? `&${query}` : "";
}

async function loadReviewImportRuns() {
  try {
    const result = await api("/api/review-imports");
    state.reviewImportRuns = Array.isArray(result.rows) ? result.rows : [];
    renderReviewImportRuns();
  } catch (e) {
    state.reviewImportRuns = [];
    setReviewImportStatus(`Failed to load import history: ${e.message}`);
  }
}

async function loadReviewSettings() {
  try {
    const result = await api("/api/review-settings");
    state.reviewSources = normalizeReviewSources(result.settings?.sources);
  } catch (e) {
    state.reviewSources = clone(DEFAULT_REVIEW_SOURCES);
    setReviewSourcesStatus(`Using default sources (${e.message}).`);
  }
}

async function loadGoogleBusinessStatus() {
  try {
    const result = await api("/api/google-business?action=status");
    state.reviewGoogle = normalizeGoogleBusinessSettings(result.google);
    setReviewGoogleStatus(googleBusinessStatusText());
  } catch (e) {
    state.reviewGoogle = normalizeGoogleBusinessSettings();
    setReviewGoogleStatus(`Google API status unavailable: ${e.message}`);
  }
}

async function loadReviewImportRun(importRunId) {
  if (!clean(importRunId)) return;
  try {
    const result = await api(`/api/review-imports?id=${encodeURIComponent(importRunId)}`);
    state.reviewImportRunId = clean(result.run?.id);
    state.reviewStagingRows = Array.isArray(result.rows) ? result.rows : [];
    if (result.run?.property_id) els.reviewsImportProperty.value = clean(result.run.property_id);
    if (result.run?.source) els.reviewsImportSource.value = clean(result.run.source);
    setReviewImportStatus(
      state.reviewStagingRows.length
        ? `Loaded ${state.reviewStagingRows.length} staged rows from ${clean(result.run?.file_name) || "saved import"}.`
        : "This import run has no staged rows."
    );
    renderReviewStaging();
  } catch (e) {
    setReviewImportStatus(`Could not load staged rows: ${e.message}`);
  }
}

async function loadLatestParsedImportRun() {
  const candidate = state.reviewImportRuns.find((run) => {
    const status = clean(run.status).toLowerCase();
    return status === "parsed" || status === "uploaded";
  });
  if (!candidate?.id) return;
  if (clean(state.reviewImportRunId) === clean(candidate.id) && state.reviewStagingRows.length) return;
  await loadReviewImportRun(candidate.id);
}

function renderReviews() {
  if (!canApp("reviews")) {
    els.reviewsCount.textContent = "0 reviews";
    els.reviewsRows.innerHTML = '<tr><td colspan="6" class="empty">Your profile has no access to Reviews.</td></tr>';
    if (els.reviewsMobileCards) {
      els.reviewsMobileCards.innerHTML = '<div class="services-mobile-empty">Your profile has no access to Reviews.</div>';
    }
    return;
  }
  setReviewScreen(state.reviewScreen);
  renderReviewPropertyOptions();
  renderReviewSettings();
  const rows = getFilteredReviews();
  renderReviewSummary(rows);
  const visibleRows = getReviewListPageRows(rows);
  renderReviewRows(visibleRows);
  renderReviewPagination(rows.length);
  syncReviewDetailModal(rows);
  renderReviewResume(rows);
  renderReviewAnalysisChart(rows);
  renderReviewQa();
}

function renderReviewSettings() {
  setReviewSettingsScreen(state.reviewSettingsScreen, false);
  renderReviewPropertyOptions();
  renderReviewProperties();
  renderReviewSources();
  if (state.reviewSettingsScreen === "config") renderGoogleBusinessSettings();
  renderReviewLastDates();
}

function renderReviewProperties() {
  if (!els.reviewsPropertiesBody) return;
  els.reviewsPropertiesBody.innerHTML = "";
  if (state.reviewProperties.length === 0) {
    els.reviewsPropertiesBody.innerHTML = '<tr><td colspan="3" class="empty">No properties yet.</td></tr>';
    return;
  }
  state.reviewProperties.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input data-review-property-name="${escape(row.id)}" value="${escape(clean(row.name))}" /></td>
      <td><input type="checkbox" data-review-property-active="${escape(row.id)}" ${row.active !== false ? "checked" : ""} /></td>
      <td class="row-actions"><button type="button" class="ghost" data-action="save-review-property" data-id="${escape(row.id)}">Save</button></td>`;
    els.reviewsPropertiesBody.appendChild(tr);
  });
}

function renderReviewSources() {
  if (!els.reviewsSourcesBody) return;
  els.reviewsSourcesBody.innerHTML = "";
  state.reviewSources.forEach((source) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(source.key)}</td>
      <td><input data-review-source-label="${escape(source.key)}" value="${escape(source.label)}" /></td>
      <td><input type="checkbox" data-review-source-active="${escape(source.key)}" ${source.active ? "checked" : ""} /></td>`;
    els.reviewsSourcesBody.appendChild(tr);
  });
}

function renderGoogleBusinessSettings() {
  if (!els.reviewsGoogleMappingsBody) return;
  setReviewGoogleStatus(state.reviewGoogle.status || googleBusinessStatusText());
  const properties = state.reviewProperties.filter((row) => row.active !== false);
  if (properties.length === 0) {
    els.reviewsGoogleMappingsBody.innerHTML = '<tr><td colspan="3" class="empty">Add a property before mapping Google locations.</td></tr>';
    return;
  }
  const locations = Array.isArray(state.reviewGoogle.locations) ? state.reviewGoogle.locations : [];
  const locationOptions = ['<option value="">Select Google location</option>']
    .concat(locations.map((location) => {
      const address = clean(location.address);
      const label = `${clean(location.title) || clean(location.reviewParent)}${address ? ` - ${address}` : ""}`;
      return `<option value="${escape(location.reviewParent)}">${escape(label)}</option>`;
    }))
    .join("");
  els.reviewsGoogleMappingsBody.innerHTML = "";
  properties.forEach((property) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(property.name)}</td>
      <td><select data-google-property-location="${escape(property.id)}">${locationOptions}</select></td>
      <td class="row-actions"><button type="button" class="ghost" data-action="sync-google-property" data-id="${escape(property.id)}" data-name="${escape(property.name)}">Sync</button></td>`;
    const select = tr.querySelector("select");
    if (select) select.value = clean(state.reviewGoogle.propertyLocations?.[property.id]);
    els.reviewsGoogleMappingsBody.appendChild(tr);
  });
}

function renderReviewLastDates() {
  if (!els.reviewsLastDatesBody) return;
  const grouped = new Map();
  const configuredSourceKeys = new Set(state.reviewSources.map((source) => clean(source.key)).filter(Boolean));
  state.reviews.forEach((row) => {
    const propertyId = clean(row.property_id);
    const source = clean(row.source);
    const date = clean(row.review_date);
    if (!configuredSourceKeys.has(source)) return;
    if (!propertyId || !source || !date) return;
    const key = `${propertyId}::${source}`;
    const current = grouped.get(key);
    if (!current || date > current.lastDate) {
      grouped.set(key, {
        propertyId,
        propertyName: clean(row.properties?.name || reviewPropertyName(propertyId) || "-"),
        source,
        lastDate: date,
      });
    }
  });
  const rows = Array.from(grouped.values()).sort((a, b) => a.lastDate.localeCompare(b.lastDate));
  els.reviewsLastDatesBody.innerHTML = "";
  if (rows.length === 0) {
    els.reviewsLastDatesBody.innerHTML = '<tr><td colspan="3" class="empty">No imported reviews yet.</td></tr>';
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escape(row.propertyName)}</td>
      <td>${escape(reviewSourceLabel(row.source))}</td>
      <td>${escape(row.lastDate)}</td>`;
    els.reviewsLastDatesBody.appendChild(tr);
  });
}

function renderReviewSourceOptions() {
  const availableSources = state.reviewSources.filter((source) => source.active);
  const currentFilter = clean(state.reviewFilters.source || els.reviewsSourceFilter.value);
  const currentImport = clean(els.reviewsImportSource.value);
  els.reviewsSourceFilter.innerHTML = ['<option value="">All</option>']
    .concat(availableSources.map((source) => `<option value="${escape(source.key)}">${escape(source.label)}</option>`))
    .join("");
  els.reviewsImportSource.innerHTML = ['<option value="">Select source</option>']
    .concat(availableSources.map((source) => `<option value="${escape(source.key)}">${escape(source.label)}</option>`))
    .join("");
  els.reviewsSourceFilter.value = currentFilter;
  if (currentImport && availableSources.some((source) => source.key === currentImport)) els.reviewsImportSource.value = currentImport;
  else els.reviewsImportSource.value = "";
}

function normalizeReviewSources(value) {
  const source = Array.isArray(value) ? value : [];
  const merged = DEFAULT_REVIEW_SOURCES.map((fallback) => {
    const match = source.find((item) => clean(item?.key) === fallback.key);
    return {
      key: fallback.key,
      label: clean(match?.label) || fallback.label,
      active: typeof match?.active === "boolean" ? match.active : fallback.active,
    };
  });
  return merged;
}

function normalizeGoogleBusinessSettings(value = {}) {
  const locations = Array.isArray(value.locations) ? value.locations : [];
  const propertyLocations = value.propertyLocations && typeof value.propertyLocations === "object" ? value.propertyLocations : {};
  return {
    connected: !!value.connected,
    connectedAt: clean(value.connectedAt),
    locationsLoadedAt: clean(value.locationsLoadedAt),
    locationsLastAttemptAt: clean(value.locationsLastAttemptAt),
    locationsLastError: clean(value.locationsLastError),
    locations: locations.map((location) => ({
      accountName: clean(location.accountName),
      locationName: clean(location.locationName),
      reviewParent: clean(location.reviewParent),
      title: clean(location.title),
      address: clean(location.address),
    })).filter((location) => location.reviewParent),
    propertyLocations,
    status: clean(value.status),
  };
}

function googleBusinessStatusText() {
  if (!state.reviewGoogle.connected) return "Google Business Profile is not connected.";
  const connectedAt = clean(state.reviewGoogle.connectedAt);
  const locationCount = state.reviewGoogle.locations.length;
  const loadedAt = clean(state.reviewGoogle.locationsLoadedAt);
  const loadedText = loadedAt ? ` Last loaded ${formatDateTimeShort(loadedAt)}.` : "";
  return `Google connected${connectedAt ? ` on ${formatDateTimeShort(connectedAt)}` : ""}. ${locationCount} location${locationCount === 1 ? "" : "s"} loaded.${loadedText}`;
}

function getFilteredReviews() {
  const { propertyId, source, search, dateFrom, dateTo, scoreFrom, scoreTo } = state.reviewFilters;
  const scoreRange = normalizeScoreRange(scoreFrom, scoreTo);
  return state.reviews.filter((row) => {
    const text = `${clean(row.title)} ${clean(row.body)} ${clean(row.reviewer_name)}`.toLowerCase();
    const score = Number(row.rating_normalized_100);
    return (!propertyId || clean(row.property_id) === propertyId) &&
      (!source || clean(row.source) === source) &&
      (!search || text.includes(search)) &&
      (!dateFrom || clean(row.review_date) >= dateFrom) &&
      (!dateTo || clean(row.review_date) <= dateTo) &&
      isScoreInRange(score, scoreRange);
  });
}

function normalizeScoreRange(fromValue, toValue) {
  const from = normalizeNumber(fromValue);
  const to = normalizeNumber(toValue);
  const hasFrom = from !== null;
  const hasTo = to !== null;
  if (!hasFrom && !hasTo) return null;
  const lower = Math.max(0, Math.min(100, hasFrom ? from : 0));
  const upper = Math.max(0, Math.min(100, hasTo ? to : 100));
  return {
    from: Math.min(lower, upper),
    to: Math.max(lower, upper),
  };
}

function isScoreInRange(score, range) {
  if (!range) return true;
  if (!Number.isFinite(score)) return false;
  return score >= range.from && score <= range.to;
}

function renderReviewSummary(rows) {
  els.reviewsCount.textContent = `${rows.length} review${rows.length === 1 ? "" : "s"}`;
  const periods = reviewPeriodBoundaries();
  els.reviewsKpiAverage12m.textContent = formatAverageOnly(averageReviewScore(rows.filter((row) => isReviewDateInRange(row.review_date, periods.last12MonthsStart, periods.today))));
  els.reviewsKpiAverageYear.textContent = formatAverageOnly(averageReviewScore(rows.filter((row) => isReviewDateInRange(row.review_date, periods.last6MonthsStart, periods.today))));
  els.reviewsKpiAverageLastMonth.textContent = formatAverageOnly(averageReviewScore(rows.filter((row) => isReviewDateInRange(row.review_date, periods.last60DaysStart, periods.today))));
  els.reviewsKpiAverageThisMonth.textContent = formatAverageOnly(averageReviewScore(rows.filter((row) => isReviewDateInRange(row.review_date, periods.last30DaysStart, periods.today))));
}

function reviewIsRecentNew(row) {
  const createdAt = clean(row?.created_at || row?.createdAt);
  if (!createdAt) return false;
  const created = new Date(createdAt);
  if (Number.isNaN(created.getTime())) return false;
  const now = Date.now();
  if (now - created.getTime() > 48 * 60 * 60 * 1000) return false;
  const updatedAt = clean(row?.updated_at || row?.updatedAt);
  if (!updatedAt) return true;
  const updated = new Date(updatedAt);
  if (Number.isNaN(updated.getTime())) return true;
  return Math.abs(updated.getTime() - created.getTime()) <= 5 * 60 * 1000;
}

function reviewDateCellHtml(row) {
  const dateText = escape(clean(row.review_date) || "-");
  const newBadge = reviewIsRecentNew(row) ? '<span class="review-new-chip">new</span>' : "";
  return `${dateText}${newBadge}`;
}

function renderReviewRows(rows) {
  els.reviewsRows.innerHTML = "";
  if (els.reviewsMobileCards) els.reviewsMobileCards.innerHTML = "";
  if (rows.length === 0) {
    els.reviewsRows.innerHTML = '<tr><td colspan="6" class="empty">No reviews match the current filters.</td></tr>';
    if (els.reviewsMobileCards) {
      els.reviewsMobileCards.innerHTML = '<div class="services-mobile-empty">No reviews match the current filters.</div>';
    }
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.dataset.reviewId = clean(row.id);
    tr.className = `clickable-row${clean(state.reviewSelectedId) === clean(row.id) ? " selected-row" : ""}`;
    const tintStyle = reviewScoreTintStyle(row.rating_normalized_100);
    tr.innerHTML = `<td>${reviewDateCellHtml(row)}</td>
      <td>${escape(clean(row.properties?.name || reviewPropertyName(row.property_id) || "-"))}</td>
      <td>${escape(clean(row.reviewer_name) || "Anonymous")}</td>
      <td>${reviewSourceIconHtml(row.source)}</td>
      <td class="message"><strong>${escape(clean(row.title) || "(no title)")}</strong><div class="review-snippet">${escape(buildReviewSnippet(row))}</div></td>
      <td>${escape(formatReviewScore(row.rating_normalized_100, row.rating_raw, row.rating_scale))}</td>`;
    if (tintStyle) tr.style.backgroundColor = tintStyle;
    els.reviewsRows.appendChild(tr);
    if (els.reviewsMobileCards) {
      const card = document.createElement("article");
      card.className = `review-mobile-card${clean(state.reviewSelectedId) === clean(row.id) ? " selected-card" : ""}`;
      card.dataset.reviewId = clean(row.id);
      card.innerHTML = `<div class="service-mobile-header">
          <div>
            <div class="service-mobile-request">${reviewDateCellHtml(row)}</div>
            <div class="service-mobile-type">${escape(clean(row.properties?.name || reviewPropertyName(row.property_id) || "-"))}</div>
          </div>
          <div class="review-mobile-score">${escape(formatReviewScore(row.rating_normalized_100, row.rating_raw, row.rating_scale))}</div>
        </div>
        <div class="review-mobile-meta">
          <span>${reviewSourceIconHtml(row.source)}</span>
          <span>${escape(clean(row.reviewer_name) || "Anonymous")}</span>
        </div>
        <div class="review-mobile-title">${escape(clean(row.title) || "(no title)")}</div>
        <div class="review-mobile-snippet">${escape(buildReviewSnippet(row))}</div>`;
      if (tintStyle) card.style.backgroundColor = tintStyle;
      els.reviewsMobileCards.appendChild(card);
    }
  });
}

function getReviewListPageRows(rows) {
  const totalPages = Math.max(1, Math.ceil(rows.length / REVIEW_LIST_PAGE_SIZE));
  state.reviewListPage = Math.min(Math.max(1, state.reviewListPage), totalPages);
  const start = (state.reviewListPage - 1) * REVIEW_LIST_PAGE_SIZE;
  return rows.slice(start, start + REVIEW_LIST_PAGE_SIZE);
}

function renderReviewPagination(totalRows) {
  if (!els.reviewsPagination) return;
  const totalPages = Math.max(1, Math.ceil(totalRows / REVIEW_LIST_PAGE_SIZE));
  state.reviewListPage = Math.min(Math.max(1, state.reviewListPage), totalPages);
  const start = totalRows === 0 ? 0 : (state.reviewListPage - 1) * REVIEW_LIST_PAGE_SIZE + 1;
  const end = Math.min(totalRows, state.reviewListPage * REVIEW_LIST_PAGE_SIZE);
  els.reviewsPagination.hidden = totalRows <= REVIEW_LIST_PAGE_SIZE;
  els.reviewsPrevPage.disabled = state.reviewListPage <= 1;
  els.reviewsNextPage.disabled = state.reviewListPage >= totalPages;
  els.reviewsPageStatus.textContent = `Page ${state.reviewListPage} of ${totalPages} · showing ${start}-${end} of ${totalRows}`;
}

function setReviewListPage(page) {
  state.reviewListPage = Math.max(1, Number(page) || 1);
  renderReviews();
}

function findReviewById(reviewId, rows = state.reviews) {
  const id = clean(reviewId);
  if (!id) return null;
  return rows.find((row) => clean(row.id) === id) || null;
}

function openReviewDetailModal() {
  if (!els.reviewDetailModal) return;
  els.reviewDetailModal.hidden = false;
  document.body.classList.add("modal-open");
}

function closeReviewDetailModal() {
  if (!els.reviewDetailModal) return;
  els.reviewDetailModal.hidden = true;
  document.body.classList.remove("modal-open");
}

function renderReviewDetail(review) {
  if (!els.reviewsDetail) return;
  if (!review) {
    els.reviewsDetail.className = "review-detail empty";
    els.reviewsDetail.textContent = "Select a review to see the full detail.";
    return;
  }
  const meta = [
    ["Date", clean(review.review_date) || "-"],
    ["Property", clean(review.properties?.name || reviewPropertyName(review.property_id) || "-")],
    ["Reviewer", clean(review.reviewer_name) || "Anonymous"],
    ["Source", reviewSourceLabel(review.source)],
    ["Score", formatReviewScore(review.rating_normalized_100, review.rating_raw, review.rating_scale)],
  ];
  const reservationValue = clean(review.source_reservation_id) || (clean(review.source).toLowerCase() === "hostelworld" ? clean(review.source_review_id) : "");
  if (reservationValue) meta.push(["Booking / Reservation", reservationValue]);
  else if (clean(review.source_review_id)) meta.push(["Source Reference", clean(review.source_review_id)]);
  const detail = [
    `<p class="review-detail-title"><strong>${escape(clean(review.title) || "(no title)")}</strong></p>`,
    `<div class="review-detail-meta">${meta.map(([label, value]) => `<div class="review-detail-meta-item"><span>${escape(label)}</span><strong>${escape(value)}</strong></div>`).join("")}</div>`,
  ];
  const subscoreHtml = renderReviewSubscores(review.subscores);
  if (subscoreHtml) detail.push(`<div class="review-detail-section"><strong>Partial scores:</strong>${subscoreHtml}</div>`);
  if (clean(review.positive_review_text)) detail.push(`<p class="review-detail-section"><strong>Positive:</strong> ${escape(clean(review.positive_review_text))}</p>`);
  if (clean(review.negative_review_text)) detail.push(`<p class="review-detail-section"><strong>Negative:</strong> ${escape(clean(review.negative_review_text))}</p>`);
  if (clean(review.body)) detail.push(`<p class="review-detail-section"><strong>Full review:</strong><br />${escape(clean(review.body))}</p>`);
  if (clean(review.host_reply_text)) detail.push(`<p class="review-detail-section"><strong>Property reply:</strong><br />${escape(clean(review.host_reply_text))}</p>`);
  els.reviewsDetail.className = "review-detail";
  els.reviewsDetail.innerHTML = detail.join("");
}

function syncReviewDetailModal(rows) {
  if (!els.reviewDetailModal || els.reviewDetailModal.hidden) return;
  const selected = findReviewById(state.reviewSelectedId, rows);
  if (!selected) {
    closeReviewDetailModal();
    return;
  }
  renderReviewDetail(selected);
}

function renderReviewResume(rows) {
  const monthly = new Map();
  const yearly = new Map();
  rows.forEach((row) => {
    const monthKey = reviewMonthKey(row.review_date);
    const yearKey = reviewYearKey(row.review_date);
    if (!monthKey || !yearKey) return;
    addReviewAggregate(monthly, monthKey, row.rating_normalized_100, row.subscores);
    addReviewAggregate(yearly, yearKey, row.rating_normalized_100, row.subscores);
  });
  const years = Array.from(yearly.values()).sort((a, b) => b.key.localeCompare(a.key));
  els.reviewsResumeRows.innerHTML = "";
  if (years.length === 0) {
    els.reviewsResumeRows.innerHTML = '<tr><td colspan="9" class="empty">No aggregate data for the current filters.</td></tr>';
    els.reviewsResumeStatus.textContent = "Grouped by year/month";
    return;
  }
  years.forEach((year) => {
    const yearAvg = year.total / year.count;
    const tr = document.createElement("tr");
    tr.className = "aggregate-total-row";
    tr.style.backgroundColor = reviewScoreTintStyle(yearAvg);
    tr.innerHTML = `<td><strong>${escape(year.key)} total</strong></td><td><strong>${escape(String(year.count))}</strong></td>${renderAggregateSubscoreCells(year, true)}<td><strong>${escape(formatAverageOnly(yearAvg))}</strong></td>`;
    els.reviewsResumeRows.appendChild(tr);
    Array.from(monthly.values())
      .filter((month) => month.key.startsWith(`${year.key}-`))
      .sort((a, b) => b.key.localeCompare(a.key))
      .forEach((month) => {
        const avg = month.total / month.count;
        const monthTr = document.createElement("tr");
        monthTr.style.backgroundColor = reviewScoreTintStyle(avg);
        monthTr.innerHTML = `<td>${escape(month.key)}</td><td>${escape(String(month.count))}</td>${renderAggregateSubscoreCells(month)}<td>${escape(formatAverageOnly(avg))}</td>`;
        els.reviewsResumeRows.appendChild(monthTr);
      });
  });
  els.reviewsResumeStatus.textContent = `${years.length} year group${years.length === 1 ? "" : "s"} shown`;
}

function renderReviewAnalysisChart(rows) {
  if (!els.reviewsAnalysisChart || !els.reviewsAnalysisLegend) return;
  const points = buildReviewAnalysisSeries(rows);
  els.reviewsAnalysisChart.innerHTML = "";
  els.reviewsAnalysisLegend.innerHTML = "";
  if (points.months.length === 0 || points.series.length === 0) {
    els.reviewsAnalysisStatus.textContent = "No scored reviews in the current filters";
    els.reviewsAnalysisChart.innerHTML = '<div class="empty">No scored reviews available for this chart.</div>';
    return;
  }

  els.reviewsAnalysisStatus.textContent = `${points.months.length} month${points.months.length === 1 ? "" : "s"} shown`;
  const width = 920;
  const height = 330;
  const margin = { top: 22, right: 24, bottom: 58, left: 46 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const scale = reviewAnalysisScale(points.series);
  const x = (index) => margin.left + (points.months.length === 1 ? plotWidth / 2 : (index / (points.months.length - 1)) * plotWidth);
  const y = (score) => margin.top + ((scale.max - score) / (scale.max - scale.min)) * plotHeight;
  const monthLabels = points.months
    .map((month, index) => ({ month, index }))
    .filter((item, index) => index === 0 || index === points.months.length - 1 || index % Math.ceil(points.months.length / 6) === 0);
  const gridLines = reviewAnalysisTicks(scale).map((score) => {
    const yy = y(score);
    return `<line x1="${margin.left}" y1="${yy}" x2="${width - margin.right}" y2="${yy}" class="analysis-grid-line" />
      <text x="${margin.left - 10}" y="${yy + 4}" text-anchor="end" class="analysis-axis-label">${score}</text>`;
  }).join("");
  const labels = monthLabels.map(({ month, index }) => {
    const xx = x(index);
    return `<text x="${xx}" y="${height - 24}" text-anchor="end" transform="rotate(-35 ${xx} ${height - 24})" class="analysis-axis-label">${escape(month)}</text>`;
  }).join("");
  const seriesSvg = points.series.map((series) => {
    const coordinates = series.points.map((point) => `${x(point.index).toFixed(1)},${y(point.average).toFixed(1)}`).join(" ");
    const circles = series.points.map((point) => {
      const month = points.months[point.index];
      const tooltip = `${series.label} | ${month} | Score ${point.average.toFixed(1)} | ${point.count} review${point.count === 1 ? "" : "s"}`;
      return `<circle class="analysis-point" cx="${x(point.index).toFixed(1)}" cy="${y(point.average).toFixed(1)}" r="${series.key === "all" ? 4 : 3}" fill="${series.color}" tabindex="0" data-tooltip="${escape(tooltip)}">
        <title>${escape(tooltip)}</title>
      </circle>`;
    }).join("");
    return `<polyline points="${coordinates}" fill="none" stroke="${series.color}" stroke-width="${series.key === "all" ? 3.2 : 2}" stroke-linecap="round" stroke-linejoin="round" opacity="${series.key === "all" ? "1" : "0.86"}" />${circles}`;
  }).join("");

  els.reviewsAnalysisChart.innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Monthly review score by source">
    <rect x="0" y="0" width="${width}" height="${height}" rx="16" class="analysis-chart-bg" />
    ${gridLines}
    <line x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${height - margin.bottom}" class="analysis-axis-line" />
    <line x1="${margin.left}" y1="${height - margin.bottom}" x2="${width - margin.right}" y2="${height - margin.bottom}" class="analysis-axis-line" />
    ${labels}
    ${seriesSvg}
  </svg><div class="analysis-tooltip" role="status" aria-live="polite"></div>`;
  els.reviewsAnalysisLegend.innerHTML = points.series.map((series) =>
    `<span class="analysis-legend-item"><span class="analysis-legend-swatch" style="background:${escape(series.color)}"></span>${escape(series.label)}</span>`
  ).join("");
  bindReviewAnalysisTooltip();
}

function bindReviewAnalysisTooltip() {
  if (!els.reviewsAnalysisChart) return;
  const tooltip = els.reviewsAnalysisChart.querySelector(".analysis-tooltip");
  if (!tooltip) return;
  const hideTooltip = () => {
    tooltip.classList.remove("visible");
    tooltip.textContent = "";
  };
  els.reviewsAnalysisChart.querySelectorAll(".analysis-point").forEach((point) => {
    const showTooltip = () => {
      const chartRect = els.reviewsAnalysisChart.getBoundingClientRect();
      const pointRect = point.getBoundingClientRect();
      tooltip.textContent = point.dataset.tooltip || "";
      tooltip.style.left = `${pointRect.left - chartRect.left + pointRect.width / 2}px`;
      tooltip.style.top = `${pointRect.top - chartRect.top}px`;
      tooltip.classList.add("visible");
    };
    point.addEventListener("mouseenter", showTooltip);
    point.addEventListener("focus", showTooltip);
    point.addEventListener("mouseleave", hideTooltip);
    point.addEventListener("blur", hideTooltip);
  });
}

function buildReviewAnalysisSeries(rows) {
  const months = Array.from(new Set(rows.map((row) => reviewMonthKey(row.review_date)).filter(Boolean))).sort();
  const monthIndex = new Map(months.map((month, index) => [month, index]));
  const sourceKeys = Array.from(new Set(rows.map((row) => clean(row.source)).filter(Boolean))).sort((a, b) => reviewSourceLabel(a).localeCompare(reviewSourceLabel(b)));
  const seriesDefinitions = [{ key: "all", label: "ALL", color: "#111827", source: "" }]
    .concat(sourceKeys.map((source, index) => ({
      key: source,
      label: reviewSourceLabel(source),
      color: REVIEW_ANALYSIS_COLORS[index % REVIEW_ANALYSIS_COLORS.length],
      source,
    })));

  const series = seriesDefinitions.map((definition) => {
    const aggregates = new Map();
    rows.forEach((row) => {
      if (definition.source && clean(row.source) !== definition.source) return;
      const month = reviewMonthKey(row.review_date);
      const score = Number(row.rating_normalized_100);
      if (!monthIndex.has(month) || !Number.isFinite(score)) return;
      const current = aggregates.get(month) || { total: 0, count: 0 };
      current.total += score;
      current.count += 1;
      aggregates.set(month, current);
    });
    return {
      ...definition,
      points: months
        .map((month, index) => {
          const item = aggregates.get(month);
          if (!item?.count) return null;
          return { index, average: item.total / item.count, count: item.count };
        })
        .filter(Boolean),
    };
  }).filter((item) => item.points.length > 0);

  return { months, series };
}

function reviewAnalysisScale(series) {
  const values = series.flatMap((item) => item.points.map((point) => point.average)).filter((value) => Number.isFinite(value));
  const minValue = values.length ? Math.min(...values) : 75;
  const min = minValue < 75 ? Math.max(0, Math.floor(minValue / 5) * 5) : 75;
  return { min, max: 100 };
}

function reviewAnalysisTicks(scale) {
  const span = scale.max - scale.min;
  const step = span <= 25 ? 5 : span <= 50 ? 10 : 20;
  const ticks = [];
  for (let value = scale.min; value <= scale.max; value += step) ticks.push(value);
  if (ticks[ticks.length - 1] !== scale.max) ticks.push(scale.max);
  return ticks;
}

function renderReviewQa() {
  if (!els.reviewsQaAnswer) return;
  if (els.reviewsQaPrompt && els.reviewsQaPrompt.value !== state.reviewQa.prompt) {
    els.reviewsQaPrompt.value = state.reviewQa.prompt;
  }
  els.reviewsQaSubmit.disabled = !!state.reviewQa.loading;
  els.reviewsQaStatus.textContent = state.reviewQa.status;
  if (!clean(state.reviewQa.answer)) {
    els.reviewsQaAnswer.className = "review-detail empty";
    els.reviewsQaAnswer.textContent = "Ask a question to analyze the filtered reviews.";
    return;
  }
  const scope = state.reviewQa.totalCount
    ? `Analyzed ${state.reviewQa.analyzedCount} of ${state.reviewQa.totalCount} filtered reviews.`
    : "";
  els.reviewsQaAnswer.className = "review-detail";
  els.reviewsQaAnswer.innerHTML = `${scope ? `<p><strong>${escape(scope)}</strong></p>` : ""}<div class="qa-answer-text">${escape(state.reviewQa.answer).replace(/\n/g, "<br />")}</div>`;
}

async function submitReviewQuestion() {
  const prompt = clean(els.reviewsQaPrompt.value);
  const filtered = getFilteredReviews();
  if (!prompt) return setReviewQaStatus("Write a question first.");
  if (filtered.length === 0) return setReviewQaStatus("There are no reviews in the current filtered scope.");
  state.reviewQa.prompt = prompt;
  state.reviewQa.loading = true;
  state.reviewQa.status = "Analyzing reviews...";
  renderReviewQa();
  try {
    const totalCount = filtered.length;
    const analyzedRows = filtered.slice(0, 250).map((row) => ({
      id: clean(row.id),
      reviewDate: clean(row.review_date),
      property: clean(row.properties?.name || reviewPropertyName(row.property_id)),
      source: reviewSourceLabel(row.source),
      reviewerName: clean(row.reviewer_name),
      ratingNormalized100: row.rating_normalized_100,
      ratingRaw: row.rating_raw,
      ratingScale: row.rating_scale,
      title: clean(row.title),
      positiveReviewText: clean(row.positive_review_text),
      negativeReviewText: clean(row.negative_review_text),
      body: clean(row.body),
      hostReplyText: clean(row.host_reply_text),
      subscores: row.subscores || {},
    }));
    const result = await api("/api/review-qa", {
      method: "POST",
      body: {
        question: prompt,
        filters: clone(state.reviewFilters),
        totalCount,
        rows: analyzedRows,
      },
    });
    state.reviewQa.answer = clean(result.answer);
    state.reviewQa.analyzedCount = Number(result.analyzedCount || analyzedRows.length);
    state.reviewQa.totalCount = Number(result.totalCount || totalCount);
    state.reviewQa.status = clean(result.note) || "Analysis complete.";
  } catch (e) {
    state.reviewQa.answer = "";
    state.reviewQa.analyzedCount = 0;
    state.reviewQa.totalCount = filtered.length;
    state.reviewQa.status = `Analysis failed: ${e.message}`;
    showToast(`Review analysis failed: ${e.message}`, "error");
  } finally {
    state.reviewQa.loading = false;
    renderReviewQa();
  }
}

function renderReviewImportRuns() {
  els.reviewsImportRuns.innerHTML = "";
  if (state.reviewImportRuns.length === 0) {
    els.reviewsImportRuns.innerHTML = '<tr><td colspan="6" class="empty">No imports yet.</td></tr>';
    return;
  }
  state.reviewImportRuns.forEach((run) => {
    const tr = document.createElement("tr");
    tr.dataset.importRunId = clean(run.id);
    tr.className = "clickable-row";
    tr.innerHTML = `<td>${escape(formatDateTimeShort(run.created_at))}</td>
      <td>${escape(clean(run.properties?.name || reviewPropertyName(run.property_id) || "-"))}</td>
      <td>${escape(reviewSourceLabel(run.source))}</td>
      <td>${escape(clean(run.file_name) || "-")}</td>
      <td>${escape(clean(run.status) || "-")}</td>
      <td>${escape(`${run.row_count_imported || 0}/${run.row_count_detected || 0}`)}</td>`;
    els.reviewsImportRuns.appendChild(tr);
  });
}

function renderReviewStaging() {
  const rows = state.reviewStagingRows;
  els.reviewsStagingCount.textContent = `${rows.length} staged row${rows.length === 1 ? "" : "s"}`;
  els.reviewsStagingRows.innerHTML = "";
  if (rows.length === 0) {
    els.reviewsStagingRows.innerHTML = '<tr><td colspan="8" class="empty">Parse a file to preview staged reviews.</td></tr>';
    return;
  }
  rows.forEach((row) => {
    const warnings = Array.isArray(row.warning_flags) ? row.warning_flags : [];
    const confidence = reviewImportConfidence(row);
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input type="checkbox" data-stage-id="${escape(row.id)}" ${row.selected_for_import ? "checked" : ""} /></td>
      <td>${escape(clean(row.review_date) || "-")}</td>
      <td>${escape(clean(row.reviewer_name) || "Anonymous")}</td>
      <td>${escape(formatReviewScore(row.rating_normalized_100, row.rating_raw, row.rating_scale))}</td>
      <td>${reviewImportConfidenceHtml(confidence)}</td>
      <td>${escape(clean(row.title) || "(no title)")}</td>
      <td class="message">${escape(buildReviewBodyPreview(row))}</td>
      <td>${warnings.length ? warnings.map((flag) => `<span class="chip status">${escape(flag)}</span>`).join(" ") : '<span class="chip status">ok</span>'}</td>`;
    els.reviewsStagingRows.appendChild(tr);
  });
}

function reviewImportConfidence(row) {
  const value = Number(row.parse_confidence);
  if (!Number.isFinite(value)) return { label: "Unknown", className: "confidence-medium", value: "" };
  const pct = Math.round(Math.max(0, Math.min(1, value)) * 100);
  if (pct >= 85) return { label: "High", className: "confidence-high", value: `${pct}%` };
  if (pct >= 65) return { label: "Check", className: "confidence-medium", value: `${pct}%` };
  return { label: "Low", className: "confidence-low", value: `${pct}%` };
}

function reviewImportConfidenceHtml(confidence) {
  return `<span class="confidence-pill ${escape(confidence.className)}"><strong>${escape(confidence.label)}</strong>${confidence.value ? `<small>${escape(confidence.value)}</small>` : ""}</span>`;
}

async function parseReviewUploads() {
  const propertyId = clean(els.reviewsImportProperty.value);
  const source = clean(els.reviewsImportSource.value);
  const uploadKind = clean(els.reviewsImportKind.value);
  const files = Array.from(els.reviewsImportFiles.files || []);
  const pastedText = clean(state.reviewImportPastedText);
  if (!propertyId || !source || !uploadKind) return setReviewImportStatus("Select a property, source, and file type first.");
  if (files.length === 0 && !pastedText) return setReviewImportStatus("Choose, drop, or paste at least one file or review text to parse.");
  setReviewImportStatus("Parsing reviews...");
  els.reviewsParseUpload.disabled = true;
  try {
    const parsedRows = [];
    if (pastedText) parsedRows.push(...parseReviewPastedText(pastedText, source));
    for (const file of files) {
      if (uploadKind === "image") {
        parsedRows.push(...(await parseReviewImageFile(file, source)));
      } else if (uploadKind === "json") {
        parsedRows.push(...(await parseReviewJsonFile(file, source)));
      } else {
        parsedRows.push(...(await parseReviewSpreadsheetFile(file, source)));
      }
    }
    if (parsedRows.length === 0) {
      setReviewImportStatus("No review rows were detected in the uploaded files.");
      state.reviewStagingRows = [];
      renderReviewStaging();
      return;
    }
    const result = await api("/api/review-imports?action=stage", {
      method: "POST",
      body: {
        propertyId,
        source,
        uploadKind,
        fileName: [files.map((file) => file.name).join(", "), pastedText ? `pasted ${reviewSourceLabel(source)} reviews` : ""].filter(Boolean).join(", "),
        fileType: [files.map((file) => file.type || inferFileType(file.name)).join(", "), pastedText ? "text/plain" : ""].filter(Boolean).join(", "),
        rows: parsedRows,
      },
    });
    state.reviewImportRunId = clean(result.run?.id);
    state.reviewStagingRows = Array.isArray(result.rows) ? result.rows : [];
    await loadReviewImportRuns();
    setReviewImportStatus(`Parsed ${state.reviewStagingRows.length} review rows. Check the preview before importing.`);
    renderReviewStaging();
  } catch (e) {
    setReviewImportStatus(`Parse failed: ${e.message}`);
    showToast(`Review parse failed: ${e.message}`, "error");
  } finally {
    els.reviewsParseUpload.disabled = false;
  }
}

async function confirmReviewImport() {
  if (!state.reviewImportRunId) return setReviewImportStatus("Parse a file first.");
  const selectedIds = state.reviewStagingRows.filter((row) => row.selected_for_import).map((row) => row.id);
  if (selectedIds.length === 0) return setReviewImportStatus("Select at least one staged row.");
  setReviewImportStatus("Importing selected reviews...");
  try {
    const result = await api("/api/review-imports?action=confirm", {
      method: "POST",
      body: { importRunId: state.reviewImportRunId, rowIds: selectedIds },
    });
    await Promise.all([loadReviews(), loadReviewImportRuns()]);
    loadSidebarReviewSummary({ silent: true }).catch(() => {});
    const replacedText = result.replacedCount ? `, replaced ${result.replacedCount} duplicate${result.replacedCount === 1 ? "" : "s"}` : "";
    const insertedText = Number.isFinite(Number(result.insertedCount)) ? ` (${result.insertedCount} new${replacedText})` : replacedText;
    setReviewImportStatus(`Imported ${result.importedCount} reviews${insertedText}.`);
    showToast(`Imported ${result.importedCount} reviews${replacedText}.`, "success");
    state.reviewStagingRows = [];
    state.reviewImportRunId = "";
    resetReviewImportForm();
    render();
  } catch (e) {
    setReviewImportStatus(`Import failed: ${e.message}`);
    showToast(`Review import failed: ${e.message}`, "error");
  }
}

function resetReviewImportForm() {
  els.reviewsImportProperty.value = "";
  els.reviewsImportSource.value = "";
  els.reviewsImportKind.value = "";
  els.reviewsImportFiles.value = "";
  state.reviewImportPastedText = "";
  renderReviewImportFileSummary();
}

function renderReviewImportFileSummary() {
  const files = Array.from(els.reviewsImportFiles.files || []);
  if (!els.reviewsImportFileSummary) return;
  const pastedText = clean(state.reviewImportPastedText);
  if (files.length === 0 && !pastedText) {
    els.reviewsImportFileSummary.textContent = "No files selected";
    return;
  }
  if (pastedText && files.length === 0) {
    const count = estimatePastedReviewCount(pastedText, els.reviewsImportSource.value);
    els.reviewsImportFileSummary.textContent = `Pasted text ready${count ? ` (${count} possible reviews)` : ""}`;
    return;
  }
  const names = files.map((file) => file.name);
  const fileSummary = files.length === 1 ? names[0] : `${files.length} files selected: ${names.join(", ")}`;
  els.reviewsImportFileSummary.textContent = pastedText ? `${fileSummary} + pasted text` : fileSummary;
}

function onReviewImportDragOver(event) {
  event.preventDefault();
  event.dataTransfer.dropEffect = "copy";
  els.reviewsImportDropzone.classList.add("drag-over");
}

function onReviewImportDragLeave(event) {
  if (event.currentTarget.contains(event.relatedTarget)) return;
  els.reviewsImportDropzone.classList.remove("drag-over");
}

function onReviewImportDrop(event) {
  event.preventDefault();
  els.reviewsImportDropzone.classList.remove("drag-over");
  setReviewImportFiles(event.dataTransfer?.files);
}

function onReviewImportPaste(event) {
  const clipboardFiles = Array.from(event.clipboardData?.files || []);
  const itemFiles = Array.from(event.clipboardData?.items || [])
    .filter((item) => item.kind === "file")
    .map((item) => item.getAsFile())
    .filter(Boolean);
  const files = clipboardFiles.length ? clipboardFiles : itemFiles;
  const pastedText = clean(event.clipboardData?.getData("text/plain"));
  if (!files.length && !pastedText) return;
  event.preventDefault();
  if (files.length) setReviewImportFiles(files);
  if (pastedText) {
    state.reviewImportPastedText = pastedText;
    renderReviewImportFileSummary();
    setReviewImportStatus("Pasted review text ready to parse.");
  }
}

function setReviewImportFiles(files) {
  const list = Array.from(files || []).filter(Boolean);
  if (!list.length) {
    setReviewImportStatus("No file was found in the drop or paste action.");
    return;
  }
  const transfer = new DataTransfer();
  list.forEach((file) => transfer.items.add(file));
  els.reviewsImportFiles.files = transfer.files;
  state.reviewImportPastedText = "";
  renderReviewImportFileSummary();
  setReviewImportStatus(`${list.length} file${list.length === 1 ? "" : "s"} ready to parse.`);
}

async function onReviewStagingToggle(event) {
  const checkbox = event.target.closest('input[data-stage-id]');
  if (!checkbox) return;
  const id = clean(checkbox.dataset.stageId);
  const row = state.reviewStagingRows.find((item) => item.id === id);
  if (!row) return;
  row.selected_for_import = !!checkbox.checked;
  try {
    const updated = await api(`/api/review-imports?id=${encodeURIComponent(id)}`, {
      method: "PUT",
      body: normalizeReviewDraft(row),
    });
    if (updated.row) Object.assign(row, updated.row);
  } catch (e) {
    checkbox.checked = !checkbox.checked;
    row.selected_for_import = checkbox.checked;
    setReviewImportStatus(`Could not update staging row: ${e.message}`);
  }
}

async function onReviewImportRunClick(event) {
  const row = event.target.closest("tr[data-import-run-id]");
  if (!row) return;
  openReviewsSettings();
  await loadReviewImportRun(clean(row.dataset.importRunId));
}

function onReviewRowClick(event) {
  const row = event.target.closest("[data-review-id]");
  if (!row) return;
  state.reviewSelectedId = clean(row.dataset.reviewId);
  const selected = findReviewById(state.reviewSelectedId, getFilteredReviews());
  renderReviews();
  renderReviewDetail(selected);
  openReviewDetailModal();
}

async function createReviewProperty() {
  try {
    const created = await api("/api/properties", {
      method: "POST",
      body: { name: `Property ${state.reviewProperties.length + 1}`, active: true },
    });
    if (created.row) {
      await loadReviewProperties();
      renderReviewSettings();
      setReviewPropertiesStatus("Property created.");
    }
  } catch (e) {
    setReviewPropertiesStatus(`Could not create property: ${e.message}`);
  }
}

async function onReviewPropertyAction(event) {
  const button = event.target.closest('button[data-action="save-review-property"]');
  if (!button) return;
  const id = clean(button.dataset.id);
  const name = clean(els.reviewsPropertiesBody.querySelector(`[data-review-property-name="${id}"]`)?.value);
  const active = !!els.reviewsPropertiesBody.querySelector(`[data-review-property-active="${id}"]`)?.checked;
  if (!name) return setReviewPropertiesStatus("Property name is required.");
  try {
    await api(`/api/properties?id=${encodeURIComponent(id)}`, {
      method: "PATCH",
      body: { name, active },
    });
    await loadReviewProperties();
    renderReviewSettings();
    setReviewPropertiesStatus("Property saved.");
  } catch (e) {
    setReviewPropertiesStatus(`Could not save property: ${e.message}`);
  }
}

async function saveReviewSettings() {
  const payload = state.reviewSources.map((source) => ({
    key: source.key,
    label: clean(els.reviewsSourcesBody.querySelector(`[data-review-source-label="${source.key}"]`)?.value) || source.label,
    active: !!els.reviewsSourcesBody.querySelector(`[data-review-source-active="${source.key}"]`)?.checked,
  }));
  try {
    await api("/api/review-settings", {
      method: "PUT",
      body: { settings: { sources: payload } },
    });
    state.reviewSources = normalizeReviewSources(payload);
    renderReviewSettings();
    setReviewSourcesStatus("Review sources saved.");
  } catch (e) {
    setReviewSourcesStatus(`Could not save review sources: ${e.message}`);
  }
}

async function connectGoogleBusiness() {
  try {
    setReviewGoogleStatus("Preparing Google connection...");
    const result = await api("/api/google-business?action=auth-url", { method: "POST", body: {} });
    if (!result.authUrl) throw new Error("Google authorization URL was not returned.");
    window.location.href = result.authUrl;
  } catch (e) {
    setReviewGoogleStatus(`Google connection failed: ${e.message}`);
    showToast(`Google connection failed: ${e.message}`, "error");
  }
}

async function loadGoogleBusinessLocations() {
  if (els.reviewsGoogleLoadLocations) els.reviewsGoogleLoadLocations.disabled = true;
  try {
    setReviewGoogleStatus("Loading Google locations...");
    const result = await api("/api/google-business?action=locations", { method: "POST", body: {} });
    state.reviewGoogle = normalizeGoogleBusinessSettings(result.google || { ...state.reviewGoogle, locations: result.locations });
    renderGoogleBusinessSettings();
    setReviewGoogleStatus(clean(result.message) || `Loaded ${state.reviewGoogle.locations.length} Google location${state.reviewGoogle.locations.length === 1 ? "" : "s"}.`);
  } catch (e) {
    setReviewGoogleStatus(`Could not load Google locations: ${e.message}`);
    showToast(`Could not load Google locations: ${e.message}`, "error");
  } finally {
    if (els.reviewsGoogleLoadLocations) els.reviewsGoogleLoadLocations.disabled = false;
  }
}

async function saveGoogleBusinessMapping() {
  if (!els.reviewsGoogleMappingsBody) return;
  const propertyLocations = {};
  els.reviewsGoogleMappingsBody.querySelectorAll("[data-google-property-location]").forEach((select) => {
    const propertyId = clean(select.dataset.googlePropertyLocation);
    const reviewParent = clean(select.value);
    if (propertyId && reviewParent) propertyLocations[propertyId] = reviewParent;
  });
  try {
    await api("/api/google-business?action=mapping", { method: "POST", body: { propertyLocations } });
    state.reviewGoogle.propertyLocations = propertyLocations;
    renderGoogleBusinessSettings();
    setReviewGoogleStatus("Google location mapping saved.");
    showToast("Google location mapping saved.", "success");
  } catch (e) {
    setReviewGoogleStatus(`Could not save Google mapping: ${e.message}`);
    showToast(`Could not save Google mapping: ${e.message}`, "error");
  }
}

async function syncGoogleBusinessReviews() {
  return syncGoogleBusinessReviewsForProperty("", "");
}

async function syncGoogleBusinessReviewsForProperty(propertyId = "", propertyName = "") {
  try {
    const name = clean(propertyName);
    const propertyIdValue = clean(propertyId);
    setReviewGoogleStatus(name ? `Syncing Google reviews for ${name}...` : "Syncing Google reviews...");
    const result = await api("/api/google-business?action=sync", { method: "POST", body: propertyIdValue ? { propertyId: propertyIdValue } : {} });
    await Promise.all([loadReviews({ useFilters: true }), loadReviewImportRuns()]);
    loadSidebarReviewSummary({ silent: true }).catch(() => {});
    render();
    const inserted = Number(result.insertedCount || 0);
    const replaced = Number(result.replacedCount || 0);
    const imported = Number(result.importedCount || 0);
    const prefix = name ? `Google sync complete for ${name}` : "Google sync complete";
    setReviewGoogleStatus(`${prefix}: ${imported} review${imported === 1 ? "" : "s"} processed, ${inserted} inserted, ${replaced} replaced.`);
    showToast(name ? `Google reviews synced for ${name}.` : "Google reviews synced.", "success");
  } catch (e) {
    const prefix = propertyName ? `Google sync failed for ${propertyName}` : "Google sync failed";
    setReviewGoogleStatus(`${prefix}: ${e.message}`);
    showToast(`${prefix}: ${e.message}`, "error");
  }
}

async function onReviewGoogleMappingAction(event) {
  const button = event.target.closest('button[data-action="sync-google-property"]');
  if (!button) return;
  const propertyId = clean(button.dataset.id);
  const propertyName = clean(button.dataset.name);
  if (!propertyId) return;
  button.disabled = true;
  try {
    await syncGoogleBusinessReviewsForProperty(propertyId, propertyName);
  } finally {
    button.disabled = false;
  }
}

async function parseReviewSpreadsheetFile(file, source) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(new Uint8Array(buffer), { type: "array" });
  const rows = [];
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    rows.push(...parseReviewSheetRows(data, source, file.name, sheetName));
  });
  return rows;
}

async function parseReviewJsonFile(file, source) {
  const text = await file.text();
  let payload;
  try {
    payload = JSON.parse(text);
  } catch {
    throw new Error(`${file.name} is not valid JSON.`);
  }
  return parseReviewJsonPayload(payload, source, file.name);
}

function parseReviewJsonPayload(payload, source, fileName) {
  const sourceKey = normalizeReviewSourceKey(source);
  if (sourceKey === "google") return parseGoogleReviewJsonPayload(payload, fileName);
  throw new Error(`JSON import is not configured for ${reviewSourceLabel(sourceKey)} yet.`);
}

function parseGoogleReviewJsonPayload(payload, fileName) {
  const rows = Array.isArray(payload?.reviews) ? payload.reviews : Array.isArray(payload) ? payload : [];
  return rows.map((row) => {
    const ratingRaw = googleStarRatingToNumber(row?.starRating);
    const body = clean(row?.comment);
    const reviewerName = clean(row?.reviewer?.displayName);
    const reviewDate = normalizeDate(clean(row?.createTime));
    const title = body ? "Google review" : "Rating only review";
    const warnings = [];
    if (!body) warnings.push("rating_only");
    if (!ratingRaw) warnings.push("missing_rating");
    if (!reviewDate) warnings.push("missing_date");
    return {
      source: "google",
      sourceReviewId: clean(row?.name),
      sourceReservationId: "",
      reviewDate,
      reviewerName,
      reviewerCountry: "",
      language: "",
      ratingRaw,
      ratingScale: 5,
      title,
      positiveReviewText: "",
      negativeReviewText: "",
      body,
      subscores: {},
      hostReplyText: clean(row?.reviewReply?.comment),
      hostReplyDate: normalizeDate(clean(row?.reviewReply?.updateTime)),
      rawText: [title, body].filter(Boolean).join("\n\n"),
      parseConfidence: 0.98,
      warningFlags: warnings,
      isValid: !!(body || ratingRaw),
      selectedForImport: true,
      rawPayload: { fileName, row },
    };
  }).filter((row) => row.isValid);
}

function googleStarRatingToNumber(value) {
  const raw = clean(value).toUpperCase();
  const map = { ONE: 1, TWO: 2, THREE: 3, FOUR: 4, FIVE: 5 };
  return map[raw] || normalizeNumber(raw);
}

function parseReviewPastedText(text, source) {
  const sourceKey = normalizeReviewSourceKey(source);
  if (sourceKey === "airbnb") return extractAirbnbReviewCandidatesFromText(text, "pasted Airbnb reviews");
  if (sourceKey === "vrbo") return extractVrboReviewCandidatesFromText(text, "pasted VRBO reviews");
  if (sourceKey === "tripadvisor") return extractTripadvisorReviewCandidatesFromText(text, "pasted Tripadvisor reviews");
  if (!window.XLSX) throw new Error("Spreadsheet parser not available.");
  const workbook = XLSX.read(text, { type: "string", raw: false });
  const rows = [];
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    rows.push(...parseReviewSheetRows(data, source, "pasted reviews", sheetName));
  });
  return rows;
}

function extractTripadvisorReviewCandidatesFromText(text, fileName, imageRatings = null) {
  const blocks = splitTripadvisorReviewBlocks(text);
  return blocks.map((block) => parseTripadvisorReviewBlock(block, fileName, imageRatings)).filter(Boolean);
}

function splitTripadvisorReviewBlocks(text) {
  const normalized = clean(text).replace(/\r/g, "\n");
  if (!normalized) return [];
  const parts = normalized
    .split(/(?=^.+\n(?:\d+\s+contributions?|contributions?)\b[\s\S]*?review\s+from\s+)/gim)
    .map((block) => clean(block))
    .filter(Boolean);
  const reviewParts = parts.filter((block) => /review\s+from\s+/i.test(block));
  return reviewParts.length ? reviewParts : (/review\s+from\s+/i.test(normalized) ? [normalized] : []);
}

function parseTripadvisorReviewBlock(block, fileName, imageRatings = null) {
  const lines = block.split(/\n/).map((line) => clean(line)).filter(Boolean);
  const dateIndex = lines.findIndex((line) => /review\s+from\s+/i.test(line));
  if (dateIndex === -1) return null;
  const reviewerName = tripadvisorReviewerName(lines, dateIndex);
  const reviewDateText = clean(lines[dateIndex].replace(/.*review\s+from\s+/i, ""));
  const reviewDate = normalizeDate(reviewDateText);
  const subscoreStart = lines.findIndex((line, index) => index > dateIndex && tripadvisorSubscoreKey(line));
  const textEnd = subscoreStart === -1 ? lines.length : subscoreStart;
  const contentLines = lines.slice(dateIndex + 1, textEnd).filter((line) => !tripadvisorNoiseLine(line));
  const title = clean(contentLines[0]) || "Tripadvisor review";
  const body = clean(contentLines.slice(1).join("\n"));
  const textRatings = tripadvisorRatingsFromText(lines);
  const ratingRaw = imageRatings?.overall || textRatings.overall || null;
  const subscores = compactSubscores({ ...(textRatings.subscores || {}), ...(imageRatings?.subscores || {}) });
  const warnings = [fileName.toLowerCase().includes("pasted") ? "tripadvisor_text" : "ocr", "tripadvisor"];
  if (!ratingRaw) warnings.push("missing_rating");
  if (!reviewDate) warnings.push("missing_date");
  if (!body) warnings.push("missing_body");

  return {
    source: "tripadvisor",
    sourceReviewId: tripadvisorSourceReviewId(reviewerName, reviewDate, title),
    sourceReservationId: "",
    reviewDate,
    reviewerName,
    reviewerCountry: "",
    language: "",
    ratingRaw,
    ratingScale: 5,
    title,
    positiveReviewText: "",
    negativeReviewText: "",
    body,
    subscores,
    hostReplyText: "",
    hostReplyDate: "",
    rawText: block,
    parseConfidence: ratingRaw && reviewDate ? 0.86 : 0.68,
    warningFlags: warnings,
    isValid: !!(title || body || ratingRaw),
    selectedForImport: true,
    rawPayload: { fileName, text: block, imageRatings },
  };
}

function tripadvisorReviewerName(lines, dateIndex) {
  const beforeDate = lines.slice(0, dateIndex).filter((line) => !tripadvisorNoiseLine(line));
  const contributionIndex = beforeDate.findIndex((line) => /\bcontributions?\b/i.test(line));
  if (contributionIndex > 0) return clean(beforeDate[contributionIndex - 1]);
  return clean(beforeDate.find((line) => /^[\w .'-]{2,60}$/i.test(line))) || "Tripadvisor reviewer";
}

function tripadvisorNoiseLine(line) {
  return /^\d+\s+contributions?$/i.test(line) ||
    /^tripadvisor$/i.test(line) ||
    /^[•.\s]+$/.test(line) ||
    /^review\s+from\s+/i.test(line);
}

function tripadvisorSubscoreKey(line) {
  const raw = clean(line).toLowerCase().replace(/[^a-z\s]/g, " ").replace(/\s+/g, " ").trim();
  const map = {
    value: "value_for_money",
    room: "rooms",
    rooms: "rooms",
    location: "location",
    cleanliness: "cleanliness",
    service: "service",
    "sleep quality": "sleep_quality",
  };
  return map[raw] || "";
}

function tripadvisorRatingsFromText(lines) {
  const result = { overall: null, subscores: {} };
  lines.forEach((line) => {
    const raw = clean(line);
    const rating = tripadvisorRatingNumberFromText(raw);
    if (!rating) return;
    const key = tripadvisorSubscoreKey(raw.replace(/\b\d(?:[.,]\d)?\s*(?:\/\s*5|of\s+5|stars?)?\b/i, ""));
    if (key) result.subscores[key] = rating;
    else if (!result.overall) result.overall = rating;
  });
  return result;
}

function tripadvisorRatingNumberFromText(value) {
  const raw = clean(value);
  const match = raw.match(/\b([1-5](?:[.,]\d)?)\s*(?:\/\s*5|of\s+5|stars?)\b/i) || raw.match(/\b([1-5])\s*green\s+(?:dots?|circles?)\b/i);
  return match ? normalizeNumber(match[1]) : null;
}

function tripadvisorSourceReviewId(reviewerName, reviewDate, title) {
  return [reviewerName, reviewDate, title]
    .map((part) => clean(part).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, ""))
    .filter(Boolean)
    .join(":");
}

async function analyzeTripadvisorGreenRatings(file) {
  try {
    const bitmap = await createImageBitmap(file);
    const maxWidth = 1600;
    const scale = Math.min(1, maxWidth / bitmap.width);
    const width = Math.max(1, Math.round(bitmap.width * scale));
    const height = Math.max(1, Math.round(bitmap.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.drawImage(bitmap, 0, 0, width, height);
    const imageData = ctx.getImageData(0, 0, width, height);
    const { data } = imageData;
    const visited = new Uint8Array(width * height);
    const components = [];
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const idx = y * width + x;
        if (visited[idx] || !isTripadvisorGreenPixel(data, idx * 4)) continue;
        const component = collectGreenComponent(data, visited, width, height, x, y);
        if (component.area >= 35 && component.width >= 5 && component.height >= 5 && component.width <= 45 && component.height <= 45) {
          components.push(scoreTripadvisorGreenComponent(imageData, component));
        }
      }
    }
    return tripadvisorRatingsFromGreenComponents(components, width);
  } catch (e) {
    return { overall: null, subscores: {} };
  }
}

function isTripadvisorGreenPixel(data, offset) {
  const r = data[offset];
  const g = data[offset + 1];
  const b = data[offset + 2];
  const a = data[offset + 3];
  return a > 120 && g > 90 && r < 70 && b < 100 && g > r * 1.5 && g > b * 1.3;
}

function collectGreenComponent(data, visited, width, height, startX, startY) {
  const stack = [[startX, startY]];
  let area = 0;
  let minX = startX;
  let maxX = startX;
  let minY = startY;
  let maxY = startY;
  let sumX = 0;
  let sumY = 0;
  while (stack.length) {
    const [x, y] = stack.pop();
    if (x < 0 || y < 0 || x >= width || y >= height) continue;
    const idx = y * width + x;
    if (visited[idx] || !isTripadvisorGreenPixel(data, idx * 4)) continue;
    visited[idx] = 1;
    area += 1;
    sumX += x;
    sumY += y;
    minX = Math.min(minX, x);
    maxX = Math.max(maxX, x);
    minY = Math.min(minY, y);
    maxY = Math.max(maxY, y);
    stack.push([x + 1, y], [x - 1, y], [x, y + 1], [x, y - 1]);
  }
  return {
    area,
    x: sumX / Math.max(1, area),
    y: sumY / Math.max(1, area),
    width: maxX - minX + 1,
    height: maxY - minY + 1,
    minX,
    maxX,
    minY,
    maxY,
  };
}

function scoreTripadvisorGreenComponent(imageData, component) {
  const { data, width, height } = imageData;
  const centerX = Math.round(component.x);
  const centerY = Math.round(component.y);
  const radius = Math.max(2, Math.floor(Math.min(component.width, component.height) * 0.28));
  let sample = 0;
  let green = 0;
  for (let y = centerY - radius; y <= centerY + radius; y += 1) {
    for (let x = centerX - radius; x <= centerX + radius; x += 1) {
      if (x < 0 || y < 0 || x >= width || y >= height) continue;
      if (((x - centerX) ** 2) + ((y - centerY) ** 2) > radius ** 2) continue;
      sample += 1;
      if (isTripadvisorGreenPixel(data, (y * width + x) * 4)) green += 1;
    }
  }
  return {
    ...component,
    fillRatio: sample ? green / sample : 0,
    filled: sample ? green / sample >= 0.45 : component.area > 120,
  };
}

function tripadvisorRatingsFromGreenComponents(components, imageWidth) {
  const groups = [];
  components.sort((a, b) => a.y - b.y || a.x - b.x).forEach((component) => {
    const group = groups.find((item) => Math.abs(item.y - component.y) <= 18);
    if (group) {
      group.items.push(component);
      group.y = group.items.reduce((sum, item) => sum + item.y, 0) / group.items.length;
    } else {
      groups.push({ y: component.y, items: [component] });
    }
  });
  const ratingRows = groups
    .map((group) => ({
      y: group.y,
      count: Math.min(5, group.items.filter((item) => item.filled).length),
      x: group.items.reduce((sum, item) => sum + item.x, 0) / group.items.length,
      items: group.items,
    }))
    .filter((group) => group.count > 0 || group.items?.length >= 3)
    .sort((a, b) => a.y - b.y);
  const overallRow = ratingRows.find((row) => row.x < imageWidth * 0.45) || ratingRows[0];
  const subscoreRows = ratingRows.filter((row) => row !== overallRow && row.x >= imageWidth * 0.45).slice(0, 6);
  const subscoreKeys = ["value_for_money", "rooms", "location", "cleanliness", "service", "sleep_quality"];
  return {
    overall: overallRow?.count || null,
    subscores: Object.fromEntries(subscoreRows.map((row, index) => [subscoreKeys[index], row.count]).filter(([key]) => key)),
  };
}

function extractVrboReviewCandidatesFromText(text, fileName) {
  const blocks = splitVrboReviewBlocks(text);
  return blocks.map((block) => parseVrboReviewBlock(block, fileName)).filter(Boolean);
}

function splitVrboReviewBlocks(text) {
  return clean(text)
    .replace(/\r/g, "\n")
    .split(/^Respond\s*$/gim)
    .map((block) => clean(block))
    .filter((block) => /Res\s*#/i.test(block) && /\bVrbo\b/i.test(block) && /Posted\s+/i.test(block));
}

function parseVrboReviewBlock(block, fileName) {
  const lines = block.split(/\n/).map((line) => clean(line)).filter(Boolean);
  const reservationIndex = lines.findIndex((line) => /^Res\s*#/i.test(line));
  const sourceIndex = lines.findIndex((line) => /^Vrbo$/i.test(line));
  const ratingIndex = lines.findIndex((line) => /^\d+(?:[.,]\d+)?\s*\/\s*10$/.test(line));
  const postedIndex = lines.findIndex((line) => /^Posted\s+/i.test(line));
  if (reservationIndex === -1 || sourceIndex === -1 || ratingIndex === -1 || postedIndex === -1) return null;

  const reviewerName = clean(lines[reservationIndex - 1]) || "VRBO guest";
  const reservationId = clean(lines[reservationIndex].replace(/^Res\s*#\s*/i, ""));
  const rating = parseReviewRating(lines[ratingIndex]);
  const reviewDate = normalizeDate(clean(lines[postedIndex].replace(/^Posted\s+/i, "")));
  const reviewLines = lines
    .slice(postedIndex + 1)
    .filter((line) => !/^Show more$/i.test(line))
    .filter((line) => !/^First picture thumbnail/i.test(line))
    .filter((line) => !/^You would not rent to /i.test(line));
  const body = clean(reviewLines.join("\n"));
  const noComment = /^This guest didn'?t leave a comment\.?$/i.test(body);
  const warnings = ["vrbo_text"];
  if (!body) warnings.push("missing_body");
  if (noComment) warnings.push("no_comment");
  if (!reviewDate) warnings.push("missing_date");
  if (!rating.raw) warnings.push("missing_rating");

  return {
    source: "vrbo",
    sourceReviewId: reservationId,
    sourceReservationId: reservationId,
    reviewDate,
    reviewerName,
    reviewerCountry: "",
    language: "",
    ratingRaw: rating.raw,
    ratingScale: rating.scale || 10,
    title: noComment ? "Rating only review" : "VRBO review",
    body,
    subscores: {},
    rawText: block,
    parseConfidence: 0.9,
    warningFlags: warnings,
    isValid: true,
    selectedForImport: true,
    rawPayload: { fileName, text: block },
  };
}

function parseReviewSheetRows(rows, source, fileName, sheetName) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const headerIndex = rows.findIndex((row) => row.some((cell) => /review|comment|guest|reviewer|rating|score/i.test(clean(cell))));
  if (headerIndex === -1) return [];
  const header = rows[headerIndex].map((cell) => clean(cell).toLowerCase());
  const col = {
    source: findHeaderIndex(header, ["brand type", "brand_type", "brand", "source", "channel"]),
    reviewDate: findHeaderIndex(header, ["review date", "review_date", "date", "submitted", "created"]),
    reviewerName: findHeaderIndex(header, ["review by", "review_by", "reviewer", "guest name", "guest", "name", "user"]),
    reviewerCountry: findHeaderIndex(header, ["country", "nationality", "origin"]),
    language: findHeaderIndex(header, ["language", "lang"]),
    ratingRaw: findHeaderIndex(header, ["review rating", "review_rating", "review score", "score", "rating", "overall"]),
    ratingScale: findHeaderIndex(header, ["scale", "out of", "max score"]),
    title: findHeaderIndex(header, ["review title", "review_title", "title", "headline", "summary"]),
    body: findHeaderIndex(header, ["review text", "review_text", "comment", "body", "text", "review"]),
    positiveReviewText: findHeaderIndex(header, ["positive review", "pros", "positive", "liked"]),
    negativeReviewText: findHeaderIndex(header, ["negative review", "cons", "negative", "disliked"]),
    hostReplyText: findHeaderIndex(header, ["review response", "review_response", "property reply", "owner reply", "host reply", "reply", "response"]),
    hostReplyDate: findHeaderIndex(header, ["review response date", "review_response_date", "response date", "reply date"]),
    sourceReviewId: findHeaderIndex(header, ["review id", "id", "reference", "ref"]),
    sourceReservationId: findHeaderIndex(header, ["reservation number", "reservation", "booking number", "confirmation number", "ref"]),
    subscoreStaff: findHeaderIndex(header, ["staff"]),
    subscoreCleanliness: findHeaderIndex(header, ["cleanliness"]),
    subscoreLocation: findHeaderIndex(header, ["location"]),
    subscoreFacilities: findHeaderIndex(header, ["facilities"]),
    subscoreComfort: findHeaderIndex(header, ["comfort"]),
    subscoreValueForMoney: findHeaderIndex(header, ["value for money", "value"]),
    subscoreRooms: findHeaderIndex(header, ["rooms", "room"]),
    subscoreService: findHeaderIndex(header, ["service"]),
    subscoreSleepQuality: findHeaderIndex(header, ["sleep quality", "sleep"]),
  };
  return rows.slice(headerIndex + 1).map((row) => {
    const brandType = clean(row[col.source]);
    const rowSource = normalizeReviewSourceKey(brandType || source);
    const parsedRating = parseReviewRating(row[col.ratingRaw]);
    const ratingRaw = parsedRating.raw;
    const explicitScale = normalizeNumber(row[col.ratingScale]) || parsedRating.scale;
    const ratingScale = explicitScale || inferRatingScale(rowSource, ratingRaw);
    const positiveReviewText = clean(row[col.positiveReviewText]);
    const negativeReviewText = clean(row[col.negativeReviewText]);
    const genericBody = clean(row[col.body]);
    const body = appendReviewBrandType(genericBody || buildCombinedReviewBody(positiveReviewText, negativeReviewText), brandType, rowSource);
    let title = clean(row[col.title]);
    const subscores = compactSubscores({
      staff: normalizeNumber(row[col.subscoreStaff]),
      cleanliness: normalizeNumber(row[col.subscoreCleanliness]),
      location: normalizeNumber(row[col.subscoreLocation]),
      facilities: normalizeNumber(row[col.subscoreFacilities]),
      comfort: normalizeNumber(row[col.subscoreComfort]),
      value_for_money: normalizeNumber(row[col.subscoreValueForMoney]),
      rooms: normalizeNumber(row[col.subscoreRooms]),
      service: normalizeNumber(row[col.subscoreService]),
      sleep_quality: normalizeNumber(row[col.subscoreSleepQuality]),
    });
    const warnings = [];
    if (!body && !title && !ratingRaw) return null;
    if (!body && !title && ratingRaw) {
      title = "Rating only review";
      warnings.push("rating_only");
    }
    if (!ratingRaw) warnings.push("missing_rating");
    if (!clean(row[col.reviewDate])) warnings.push("missing_date");
    if (positiveReviewText && negativeReviewText) warnings.push("split_review");
    const sourceReviewId = clean(row[col.sourceReviewId]);
    const sourceReservationId = clean(row[col.sourceReservationId]) || (clean(source).toLowerCase() === "hostelworld" ? sourceReviewId : "");
    return {
      source: rowSource,
      sourceReviewId,
      sourceReservationId,
      reviewDate: normalizeDate(clean(row[col.reviewDate])),
      reviewerName: clean(row[col.reviewerName]),
      reviewerCountry: clean(row[col.reviewerCountry]),
      language: clean(row[col.language]),
      ratingRaw,
      ratingScale,
      title,
      positiveReviewText,
      negativeReviewText,
      body,
      subscores,
      hostReplyText: clean(row[col.hostReplyText]),
      hostReplyDate: normalizeDate(clean(row[col.hostReplyDate])),
      rawText: [title, body].filter(Boolean).join("\n\n"),
      parseConfidence: 0.95,
      warningFlags: warnings,
      isValid: !!(body || title),
      selectedForImport: true,
      rawPayload: { fileName, sheetName, brandType, row },
    };
  }).filter(Boolean);
}

async function parseReviewImageFile(file, source) {
  if (!window.Tesseract) throw new Error("OCR library not available.");
  const ocrLanguage = reviewOcrLanguages(source);
  const { data } = await window.Tesseract.recognize(file, ocrLanguage);
  const text = clean(data?.text);
  if (!text) return [];
  if (normalizeReviewSourceKey(source) === "tripadvisor") {
    const ratings = await analyzeTripadvisorGreenRatings(file);
    const tripadvisorRows = extractTripadvisorReviewCandidatesFromText(text, file.name, ratings);
    if (tripadvisorRows.length) return tripadvisorRows;
  }
  return extractReviewCandidatesFromText(text, source, file.name);
}

function reviewOcrLanguages(source) {
  if (normalizeReviewSourceKey(source) !== "airbnb") return "eng";
  return "eng+por+spa+fra+ita+deu+nld";
}

function estimatePastedReviewCount(text, source) {
  const sourceKey = normalizeReviewSourceKey(source);
  if (sourceKey === "vrbo") return splitVrboReviewBlocks(text).length;
  if (sourceKey === "airbnb") return estimateAirbnbCardCount(text);
  if (sourceKey === "tripadvisor") return splitTripadvisorReviewBlocks(text).length;
  return 0;
}

function estimateAirbnbCardCount(text) {
  const detailCount = (clean(text).match(/ver detalhes/gi) || []).length;
  if (detailCount) return detailCount;
  return splitAirbnbReviewBlocks(text).length;
}

function extractReviewCandidatesFromText(text, source, fileName) {
  if (normalizeReviewSourceKey(source) === "tripadvisor") {
    const tripadvisorRows = extractTripadvisorReviewCandidatesFromText(text, fileName);
    if (tripadvisorRows.length) return tripadvisorRows;
  }
  if (normalizeReviewSourceKey(source) === "airbnb") {
    const airbnbRows = extractAirbnbReviewCandidatesFromText(text, fileName);
    if (airbnbRows.length) return airbnbRows;
  }
  const blocks = text
    .split(/\n{2,}/)
    .map((block) => clean(block))
    .filter(Boolean);
  const rows = [];
  blocks.forEach((block) => {
    const lines = block.split(/\n/).map((line) => clean(line)).filter(Boolean);
    if (lines.length < 2) return;
    const ratingMatch = block.match(/(\d+(?:[.,]\d+)?)\s*(?:\/|out of )\s*(5|10)/i) || block.match(/(\d(?:[.,]\d)?)\s*stars?/i);
    const dateMatch = block.match(/\b(\d{4}-\d{2}-\d{2}|\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4})\b/);
    const title = lines[0];
    const body = lines.slice(1).join(" ");
    if (!body && !title) return;
    const ratingRaw = normalizeNumber(ratingMatch?.[1]);
    const ratingScale = normalizeNumber(ratingMatch?.[2]) || inferRatingScale(source, ratingRaw) || 5;
    const warnings = ["ocr"];
    if (!ratingRaw) warnings.push("missing_rating");
    if (!dateMatch) warnings.push("missing_date");
    rows.push({
      source,
      reviewDate: normalizeDate(dateMatch?.[1]),
      reviewerName: "",
      ratingRaw,
      ratingScale,
      title,
      body,
      rawText: block,
      parseConfidence: 0.55,
      warningFlags: warnings,
      isValid: true,
      selectedForImport: true,
      rawPayload: { fileName, text: block },
    });
  });
  return rows;
}

function extractAirbnbReviewCandidatesFromText(text, fileName) {
  const normalizedText = clean(text).replace(/\r/g, "\n");
  const blocks = splitAirbnbReviewBlocks(normalizedText, fileName.includes("pasted"))
    .map((block) => clean(block))
    .filter(Boolean);
  const rows = [];
  blocks.forEach((block) => {
    const lines = block.split(/\n/).map((line) => clean(line)).filter(Boolean);
    if (lines.length < 3) return;
    const dateLineIndex = lines.findIndex(isAirbnbDateLine);
    const ratingLineIndex = lines.findIndex(isAirbnbRatingLine);
    if (dateLineIndex === -1 || ratingLineIndex === -1) return;
    const rawReviewerName = normalizeAirbnbReviewerName(lines.slice(0, dateLineIndex));
    const dateLine = lines[dateLineIndex];
    const ratingData = extractAirbnbRatingAndBody(lines, ratingLineIndex);
    const body = ratingData.body;
    const ratingRaw = ratingData.ratingRaw || 5;
    const warnings = fileName.includes("pasted") ? ["airbnb_text"] : ["ocr", "airbnb_screenshot"];
    if (!body) warnings.push("missing_body");
    const hasReadableText = hasLatinText(rawReviewerName) || hasLatinText(body);
    if (!hasReadableText) warnings.push("non_latin_notparsed");
    rows.push({
      source: "airbnb",
      reviewDate: normalizeDate(dateLine),
      reviewerName: hasReadableText ? (rawReviewerName || "Airbnb guest") : "notparsed",
      ratingRaw,
      ratingScale: 5,
      title: "Airbnb review",
      body: hasReadableText ? body : "",
      rawText: block,
      parseConfidence: 0.72,
      warningFlags: warnings,
      isValid: true,
      selectedForImport: true,
      rawPayload: { fileName, text: block, dateLine },
    });
  });
  return rows;
}

function normalizeAirbnbReviewerName(lines) {
  const names = lines.map((line) => clean(line)).filter(Boolean);
  const unique = [];
  names.forEach((name) => {
    if (!unique.some((item) => item.toLowerCase() === name.toLowerCase())) unique.push(name);
  });
  return clean(unique.join(" "));
}

function extractAirbnbRatingAndBody(lines, ratingLineIndex) {
  const ratingLine = clean(lines[ratingLineIndex]);
  let ratingRaw = normalizeNumber(ratingLine.match(/(\d+(?:[.,]\d+)?)\s*$/)?.[1]) || null;
  let bodyStartIndex = ratingLineIndex + 1;
  for (let i = ratingLineIndex + 1; i < Math.min(lines.length, ratingLineIndex + 5); i += 1) {
    const line = clean(lines[i]);
    const numberMatch = line.match(/(\d+(?:[.,]\d+)?)/);
    if (!ratingRaw && numberMatch) ratingRaw = normalizeNumber(numberMatch[1]);
    if (/avalia[cç][aã]o|rating|^(\d+(?:[.,]\d+)?)$|^\d+\s+de\s+\d+/i.test(line)) {
      bodyStartIndex = i + 1;
      continue;
    }
    break;
  }
  return {
    ratingRaw,
    body: clean(lines.slice(bodyStartIndex).filter((line) => !/^ver detalhes$/i.test(line)).join(" ")),
  };
}

function hasLatinText(value) {
  return /[a-z]/i.test(clean(value));
}

function splitAirbnbReviewBlocks(text, strictDetails = false) {
  if (strictDetails) {
    return text.split(/ver detalhes/i).map((block) => clean(block)).filter(Boolean);
  }
  const lines = text.split(/\n/);
  const blocks = [];
  let current = [];
  lines.forEach((line) => {
    if (isAirbnbDateLine(line) && current.some((item) => /^ver detalhes$/i.test(clean(item)))) {
      blocks.push(current.join("\n"));
      current = [];
    }
    current.push(line);
  });
  if (current.length) blocks.push(current.join("\n"));
  return blocks;
}

function isAirbnbDateLine(line) {
  const raw = clean(line);
  const separator = "[\\s\\u2009\\u202f]*[^\\da-z.]+[\\s\\u2009\\u202f]*";
  return new RegExp(`\\b\\d{1,2}${separator}\\d{1,2}\\s+(?:de\\s+)?[a-zç.]+\\s+de\\s+\\d{4}\\b`, "i").test(raw) ||
    new RegExp(`\\b\\d{1,2}\\s+de\\s+[a-zç.]+${separator}\\d{1,2}\\s+de\\s+[a-zç.]+\\s+de\\s+\\d{4}\\b`, "i").test(raw) ||
    new RegExp(`\\b\\d{1,2}\\s+de\\s+[a-zç.]+\\s+de\\s+\\d{4}${separator}\\d{1,2}\\s+de\\s+[a-zç.]+\\s+de\\s+\\d{4}\\b`, "i").test(raw);
}

function isAirbnbRatingLine(line) {
  const raw = clean(line).toLowerCase();
  return /qualidade\s+geral/.test(raw) || /general\s+quality/.test(raw) || /[\u2605*]\s*5\b/.test(raw) || /^.{0,24}\b5$/.test(raw);
}

function normalizeReviewDraft(row) {
  return {
    propertyId: clean(row.property_id),
    source: clean(row.source),
    sourceReviewId: clean(row.source_review_id),
    sourceReservationId: clean(row.source_reservation_id),
    reviewDate: clean(row.review_date),
    reviewerName: clean(row.reviewer_name),
    reviewerCountry: clean(row.reviewer_country),
    language: clean(row.language),
    ratingRaw: row.rating_raw,
    ratingScale: row.rating_scale,
    ratingNormalized100: row.rating_normalized_100,
    title: clean(row.title),
    positiveReviewText: clean(row.positive_review_text),
    negativeReviewText: clean(row.negative_review_text),
    body: clean(row.body),
    subscores: row.subscores || {},
    hostReplyText: clean(row.host_reply_text),
    hostReplyDate: clean(row.host_reply_date),
    rawText: clean(row.raw_text),
    parseConfidence: row.parse_confidence,
    warningFlags: row.warning_flags || [],
    rawPayload: row.raw_payload || {},
    selectedForImport: !!row.selected_for_import,
    isValid: !!row.is_valid,
  };
}

function reviewPropertyName(propertyId) {
  return state.reviewProperties.find((row) => clean(row.id) === clean(propertyId))?.name || "";
}

function reviewSourceLabel(source) {
  const raw = clean(source).toLowerCase();
  const configured = state.reviewSources.find((item) => item.key === raw);
  if (configured?.label) return configured.label;
  if (raw === "booking") return "Booking.com";
  if (raw === "hostelworld") return "Hostelworld";
  if (raw === "expedia") return "Expedia";
  if (raw === "hotels") return "Expedia";
  if (raw === "airbnb") return "Airbnb";
  if (raw === "vrbo") return "VRBO";
  if (raw === "tripadvisor") return "Tripadvisor";
  if (raw === "google") return "Google";
  return raw || "-";
}

function reviewSourceIconHtml(source) {
  const key = clean(source).toLowerCase();
  const label = reviewSourceLabel(key);
  const domain = reviewSourceIconDomain(key);
  const initials = sourceInitials(label);
  const title = escape(label);
  if (!domain) {
    return `<span class="source-icon source-icon-fallback" title="${title}" aria-label="${title}">${escape(initials)}</span>`;
  }
  const iconUrl = `https://www.google.com/s2/favicons?domain=${encodeURIComponent(domain)}&sz=64`;
  return `<span class="source-icon" title="${title}" aria-label="${title}">
    <img src="${escape(iconUrl)}" alt="" loading="lazy" referrerpolicy="no-referrer" onerror="this.parentElement.classList.add('source-icon-fallback'); this.remove();" />
    <span class="source-icon-text">${escape(initials)}</span>
  </span>`;
}

function reviewSourceIconDomain(source) {
  const raw = clean(source).toLowerCase();
  if (raw === "booking") return "booking.com";
  if (raw === "hostelworld") return "hostelworld.com";
  if (raw === "expedia" || raw === "hotels") return "expedia.com";
  if (raw === "airbnb") return "airbnb.com";
  if (raw === "vrbo") return "vrbo.com";
  if (raw === "tripadvisor") return "tripadvisor.com";
  if (raw === "google") return "google.com";
  return "";
}

function sourceInitials(label) {
  const words = clean(label).replace(".com", "").split(/\s+/).filter(Boolean);
  if (words.length === 0) return "?";
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
  return words.slice(0, 2).map((word) => word[0]).join("").toUpperCase();
}

function buildReviewBodyPreview(row) {
  const body = clean(row.body);
  if (body) return body;
  return buildCombinedReviewBody(clean(row.positive_review_text), clean(row.negative_review_text)) || "-";
}

function formatReviewScore(normalized, raw, scale) {
  if (normalized || normalized === 0) {
    const parts = [`${Number(normalized).toFixed(0)}/100`];
    if (raw && scale) parts.push(`(${raw}/${scale})`);
    return parts.join(" ");
  }
  if (raw && scale) return `${raw}/${scale}`;
  if (raw) return `${raw}`;
  return "-";
}

function formatMoney(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return "0,00 €";
  const formatted = new Intl.NumberFormat("de-DE", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(num);
  return `${formatted} €`;
}

function formatDateOnly(value) {
  const raw = clean(value);
  if (!raw) return "-";
  return raw.slice(0, 10);
}

function formatDateInLisbon(value) {
  const raw = clean(value);
  if (!raw) return "";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return raw.slice(0, 10);
  return new Intl.DateTimeFormat("sv-SE", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
}

function findHeaderIndex(header, candidates) {
  const normalizedHeader = header.map(normalizeHeaderCell);
  const normalizedCandidates = candidates.map(normalizeHeaderCell);
  const exactIndex = normalizedHeader.findIndex((cell) => normalizedCandidates.some((candidate) => cell === candidate));
  if (exactIndex !== -1) return exactIndex;
  return normalizedHeader.findIndex((cell) => normalizedCandidates.some((candidate) => cell.includes(candidate)));
}

function normalizeHeaderCell(value) {
  return clean(value).toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
}

function normalizeNumber(value) {
  const raw = clean(value).replace(",", ".");
  if (!raw) return null;
  const num = Number(raw);
  return Number.isFinite(num) ? num : null;
}

function parseReviewRating(value) {
  const raw = clean(value);
  const numeric = normalizeNumber(raw);
  if (numeric !== null) return { raw: numeric, scale: null };
  const match = raw.replace(",", ".").match(/(\d+(?:\.\d+)?)\s*(?:\/|out of)\s*(\d+(?:\.\d+)?)/i);
  if (!match) return { raw: null, scale: null };
  return {
    raw: normalizeNumber(match[1]),
    scale: normalizeNumber(match[2]),
  };
}

function normalizeReviewSourceKey(value) {
  const raw = clean(value).toLowerCase();
  if (raw.includes("booking")) return "booking";
  if (raw.includes("hostelworld")) return "hostelworld";
  if (raw.includes("expedia") || raw.includes("hotel")) return "expedia";
  if (raw.includes("airbnb")) return "airbnb";
  if (raw.includes("vrbo") || raw.includes("homeaway")) return "vrbo";
  if (raw.includes("trip")) return "tripadvisor";
  if (raw.includes("google")) return "google";
  return raw || "unknown";
}

function inferRatingScale(source, ratingRaw) {
  const raw = clean(source).toLowerCase();
  if (raw === "booking") return 10;
  if (raw === "hostelworld") return 10;
  if (raw === "hotels") return 10;
  if (raw === "airbnb") return ratingRaw && ratingRaw > 5 ? 10 : 5;
  if (raw === "vrbo") return 10;
  if (raw === "google" || raw === "tripadvisor") return 5;
  if (raw === "expedia") return ratingRaw && ratingRaw > 5 ? 10 : 5;
  return ratingRaw && ratingRaw > 5 ? 10 : 5;
}

function inferFileType(fileName) {
  const lower = clean(fileName).toLowerCase();
  if (lower.endsWith(".json")) return "application/json";
  if (lower.endsWith(".csv")) return "text/csv";
  if (lower.endsWith(".xlsx")) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (lower.endsWith(".xls")) return "application/vnd.ms-excel";
  return "";
}

function buildCombinedReviewBody(positiveText, negativeText) {
  const parts = [];
  if (clean(positiveText)) parts.push(`Positive: ${clean(positiveText)}`);
  if (clean(negativeText)) parts.push(`Negative: ${clean(negativeText)}`);
  return parts.join("\n");
}

function appendReviewBrandType(body, brandType, source) {
  const base = clean(body);
  const brand = clean(brandType);
  if (!brand || normalizeReviewSourceKey(source) !== "expedia") return base;
  const suffix = `Brand type: ${brand}`;
  if (!base) return suffix;
  if (base.toLowerCase().includes("brand type:")) return base;
  return `${base}\n\n${suffix}`;
}

function compactSubscores(value) {
  return Object.fromEntries(
    Object.entries(value || {}).filter(([, item]) => item !== null && item !== undefined && item !== "")
  );
}

function renderReviewSubscores(subscores) {
  const entries = Object.entries(subscores || {}).filter(([, value]) => value !== null && value !== undefined && value !== "");
  if (entries.length === 0) return "";
  const items = entries
    .map(([key, value]) => `<span class="chip status">${escape(formatSubscoreKey(key))}: ${escape(String(value))}</span>`)
    .join(" ");
  return `<div class="review-subscores">${items}</div>`;
}

function formatSubscoreKey(key) {
  return clean(key)
    .replaceAll("_", " ")
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function reviewScoreTintStyle(normalized) {
  const score = Number(normalized);
  if (!Number.isFinite(score)) return "";
  if (score >= 100) return "rgba(22, 163, 74, 0.24)";
  if (score >= 90) return "rgba(46, 159, 66, 0.14)";
  if (score >= 75) return "rgba(210, 171, 35, 0.14)";
  if (score >= 60) return "rgba(224, 138, 58, 0.14)";
  return "rgba(212, 76, 76, 0.14)";
}

function buildReviewSnippet(row) {
  const text = buildReviewBodyPreview(row);
  return text.length > 180 ? `${text.slice(0, 177)}...` : text;
}

function reviewMonthKey(value) {
  const raw = clean(value);
  if (!raw) return "";
  return /^\d{4}-\d{2}/.test(raw) ? raw.slice(0, 7) : "";
}

function reviewYearKey(value) {
  const raw = clean(value);
  if (!raw) return "";
  return /^\d{4}/.test(raw) ? raw.slice(0, 4) : "";
}

function addReviewAggregate(map, key, score, subscores = {}) {
  const value = Number(score);
  if (!Number.isFinite(value)) return;
  const item = map.get(key) || { key, count: 0, total: 0, subscores: {} };
  item.count += 1;
  item.total += value;
  REVIEW_SUBSCORE_KEYS.forEach((subscoreKey) => {
    const subscore = Number(subscores?.[subscoreKey]);
    if (!Number.isFinite(subscore)) return;
    const current = item.subscores[subscoreKey] || { count: 0, total: 0 };
    current.count += 1;
    current.total += subscore;
    item.subscores[subscoreKey] = current;
  });
  map.set(key, item);
}

function renderAggregateSubscoreCells(aggregate, strong = false) {
  return REVIEW_SUBSCORE_KEYS.map((key) => {
    const item = aggregate.subscores?.[key];
    const value = item?.count ? formatAverageOnly(item.total / item.count) : "-";
    return strong ? `<td class="subscore-cell"><strong>${escape(value)}</strong></td>` : `<td class="subscore-cell">${escape(value)}</td>`;
  }).join("");
}

function averageReviewScore(rows) {
  const values = rows.map((row) => Number(row.rating_normalized_100)).filter((value) => Number.isFinite(value));
  if (values.length === 0) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function formatAverageOnly(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num.toFixed(1) : "-";
}

function renderSidebarReviewSummary() {
  if (!els.sidebarReviewSummaryCard || !els.sidebarReviewSummaryBody || !els.sidebarReviewSummaryStatus) return;
  if (!canApp("communications")) {
    els.sidebarReviewSummaryCard.hidden = true;
    return;
  }
  els.sidebarReviewSummaryCard.hidden = false;
  const summary = state.sidebarReviewSummary;
  if (!summary) {
    els.sidebarReviewSummaryStatus.textContent = state.sidebarReviewSummaryLoaded ? "No review data" : "Loading...";
    els.sidebarReviewSummaryBody.innerHTML = '<p class="sidebar-summary-empty">Average ratings for Hostel and Cruz will appear here.</p>';
    return;
  }
  els.sidebarReviewSummaryStatus.textContent = "";
  const currentLabel = summary.months?.currentLabel || "Current";
  const previousLabel = summary.months?.previousLabel || "Past";
  const hostel = summary.properties?.hostel || {};
  const cruz = summary.properties?.cruz || {};
  const overall = summary.properties?.overall || {};
  els.sidebarReviewSummaryBody.innerHTML = `
    <div class="sidebar-summary-grid">
      <div class="sidebar-summary-grid-head">
        <span class="sidebar-summary-head-label">Property</span>
        <span class="sidebar-summary-head-value">${escape(currentLabel)}</span>
        <span class="sidebar-summary-head-value">${escape(previousLabel)}</span>
      </div>
      <div class="sidebar-summary-grid-row">
        <span class="sidebar-summary-label">Hostel</span>
        <span class="sidebar-summary-value">${escape(formatAverageOnly(hostel.currentAverage))}</span>
        <span class="sidebar-summary-value">${escape(formatAverageOnly(hostel.previousAverage))}</span>
      </div>
      <div class="sidebar-summary-grid-row">
        <span class="sidebar-summary-label">Cruz</span>
        <span class="sidebar-summary-value">${escape(formatAverageOnly(cruz.currentAverage))}</span>
        <span class="sidebar-summary-value">${escape(formatAverageOnly(cruz.previousAverage))}</span>
      </div>
      <div class="sidebar-summary-grid-row sidebar-summary-grid-row-total">
        <span class="sidebar-summary-label">Overall</span>
        <span class="sidebar-summary-value">${escape(formatAverageOnly(overall.currentAverage))}</span>
        <span class="sidebar-summary-value">${escape(formatAverageOnly(overall.previousAverage))}</span>
      </div>
    </div>`;
}

async function loadSidebarReviewSummary({ silent = false } = {}) {
  if (!canApp("communications")) return;
  try {
    const result = await api("/api/review-sidebar-summary");
    state.sidebarReviewSummary = result.summary || null;
    state.sidebarReviewSummaryLoaded = true;
    if (!silent && !state.sidebarReviewSummary) showToast("No review summary data available.", "info");
  } catch (e) {
    state.sidebarReviewSummary = null;
    state.sidebarReviewSummaryLoaded = true;
    if (!silent) showToast(`Could not load review snapshot: ${e.message}`, "error");
  }
  renderSidebarReviewSummary();
}

function reviewPeriodBoundaries() {
  const today = new Date();
  const last12MonthsStart = shiftReviewDate(today, { months: -12 });
  const last6MonthsStart = shiftReviewDate(today, { months: -6 });
  const last60DaysStart = shiftReviewDate(today, { days: -60 });
  const last30DaysStart = shiftReviewDate(today, { days: -30 });
  return {
    today: formatDate(today),
    last12MonthsStart: formatDate(last12MonthsStart),
    last6MonthsStart: formatDate(last6MonthsStart),
    last60DaysStart: formatDate(last60DaysStart),
    last30DaysStart: formatDate(last30DaysStart),
  };
}

function shiftReviewDate(date, { months = 0, days = 0 } = {}) {
  const result = new Date(date);
  if (months) result.setMonth(result.getMonth() + months);
  if (days) result.setDate(result.getDate() + days);
  return result;
}

function isReviewDateInRange(value, start, end) {
  const raw = clean(value);
  if (!raw) return false;
  return (!start || raw >= start) && (!end || raw <= end);
}

function setReviewsStatus(text) {
  els.reviewsStatus.textContent = text;
}

function isErrorStatusMessage(text) {
  const raw = clean(text).toLowerCase();
  if (!raw) return false;
  return [
    "failed",
    "could not",
    "required",
    "must",
    "invalid",
    "error",
    "no ",
  ].some((token) => raw.includes(token));
}

function setGroupsStatus(text) {
  const isError = isErrorStatusMessage(text);
  if (els.groupsStatus) {
    els.groupsStatus.textContent = text;
    els.groupsStatus.classList.toggle("status-error", isError);
  }
  if (els.groupsStatusFooter) {
    els.groupsStatusFooter.textContent = text;
    els.groupsStatusFooter.classList.toggle("status-error", isError);
  }
}

function setGroupsSettingsStatus(text) {
  if (els.groupsSettingsStatus) els.groupsSettingsStatus.textContent = text;
}

function setServicesStatus(text) {
  if (els.servicesStatus) els.servicesStatus.textContent = text;
}

function setServicesDbStatus(text) {
  if (els.servicesDbStatus) els.servicesDbStatus.textContent = text;
}

function setServicesSettingsStatus(text) {
  if (els.servicesSettingsStatus) els.servicesSettingsStatus.textContent = text;
}

function setReviewImportStatus(text) {
  els.reviewsImportStatus.textContent = text;
}

function setReviewPropertiesStatus(text) {
  if (els.reviewsPropertiesStatus) els.reviewsPropertiesStatus.textContent = text;
}

function setReviewSourcesStatus(text) {
  if (els.reviewsSourcesStatus) els.reviewsSourcesStatus.textContent = text;
}

function setReviewGoogleStatus(text) {
  state.reviewGoogle.status = text;
  if (els.reviewsGoogleStatus) els.reviewsGoogleStatus.textContent = text;
}

function setReviewQaStatus(text) {
  state.reviewQa.status = text;
  renderReviewQa();
}

function clean(value) {
  return String(value ?? "").trim();
}

function escape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

