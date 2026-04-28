const SHEET_NAME = "Comunicações";
const REVIEW_LIST_PAGE_SIZE = 100;
const REVIEW_FETCH_PAGE_SIZE = 1500;
const AUTO_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const AUTO_REFRESH_ON_FOCUS_AFTER_MS = 60 * 1000;
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
const APP_FEATURE_OPTIONS = ["communications", "lost-found", "reviews", "groups", "services"];
const SETTINGS_FEATURE_OPTIONS = ["communications", "reviews", "groups", "services", "admin-users"];

const DEFAULT_SETTINGS = {
  communications: {
    categories: [
      { name: "Warning", color: "#ffd89b" },
      { name: "Maintenance", color: "#a9f0df" },
      { name: "Information", color: "#add4ff" },
      { name: "very important", color: "#ffb3c2" },
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
  groupSort: { key: "dates", dir: "asc" },
  groupSettingsTab: "config",
  groupSettings: clone(DEFAULT_GROUP_SETTINGS),
  groupsShowActive: true,
  services: [],
  servicesLoaded: false,
  serviceSettings: clone(DEFAULT_SERVICE_SETTINGS),
  serviceSettingsLoaded: false,
  serviceProviders: [],
  serviceFilters: { showActive: true, createdFrom: "", createdTo: "", dateFrom: "", dateTo: "", name: "" },
  serviceDraft: emptyServiceDraft(),
  serviceSelectedId: "",
  serviceFlightStatuses: {
    cache: {},
    timer: null,
    sequence: 0,
    initialized: false,
  },
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
  settingsSection: "communications",
  autoRefreshTimer: null,
  lastAutoRefreshAt: 0,
  autoRefreshRunning: false,
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
  topbar: document.querySelector(".topbar"),
  navCommunications: document.getElementById("nav-communications"),
  navLostFound: document.getElementById("nav-lost-found"),
  navReviews: document.getElementById("nav-reviews"),
  navGroups: document.getElementById("nav-groups"),
  navServices: document.getElementById("nav-services"),
  openSettings: document.getElementById("open-settings"),
  closeSettings: document.getElementById("close-settings"),
  viewCommunications: document.getElementById("view-communications"),
  viewLostFound: document.getElementById("view-lost-found"),
  viewReviews: document.getElementById("view-reviews"),
  viewServices: document.getElementById("view-services"),
  viewSettings: document.getElementById("view-settings"),
  settingsMenuCommunications: document.getElementById("settings-menu-communications"),
  settingsMenuReviews: document.getElementById("settings-menu-reviews"),
  settingsMenuGroups: document.getElementById("settings-menu-groups"),
  settingsMenuServices: document.getElementById("settings-menu-services"),
  settingsMenuAdminUsers: document.getElementById("settings-menu-admin-users"),
  settingsViewCommunications: document.getElementById("settings-view-communications"),
  settingsViewReviews: document.getElementById("settings-view-reviews"),
  settingsViewGroups: document.getElementById("settings-view-groups"),
  settingsViewServices: document.getElementById("settings-view-services"),
  settingsViewAdminUsers: document.getElementById("settings-view-admin-users"),
  settingsReviewsImportTab: document.getElementById("settings-reviews-import-tab"),
  settingsReviewsConfigTab: document.getElementById("settings-reviews-config-tab"),
  settingsReviewsImportPanel: document.getElementById("settings-reviews-import-panel"),
  settingsReviewsConfigPanel: document.getElementById("settings-reviews-config-panel"),
  closeSettingsAdmin: document.getElementById("close-settings-admin"),
  closeSettingsReviews: document.getElementById("close-settings-reviews"),
  closeSettingsGroups: document.getElementById("close-settings-groups"),
  closeSettingsServices: document.getElementById("close-settings-services"),
  adminUserEmail: document.getElementById("admin-user-email"),
  adminUserPassword: document.getElementById("admin-user-password"),
  adminUserProfile: document.getElementById("admin-user-profile"),
  adminCreateUser: document.getElementById("admin-create-user"),
  adminRefreshUsers: document.getElementById("admin-refresh-users"),
  adminUsersStatus: document.getElementById("admin-users-status"),
  adminUsersBody: document.getElementById("admin-users-body"),
  profilesBody: document.getElementById("profiles-body"),
  addProfile: document.getElementById("add-profile"),
  profilesStatus: document.getElementById("profiles-status"),
  rows: document.getElementById("rows"),
  lostFoundRows: document.getElementById("lost-found-rows"),
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
  search: document.getElementById("search"),
  showActive: document.getElementById("show-active"),
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
  viewGroups: document.getElementById("view-groups"),
  groupsNew: document.getElementById("groups-new"),
  groupsStatus: document.getElementById("groups-status"),
  groupReservationNumber: document.getElementById("group-reservation-number"),
  groupName: document.getElementById("group-name"),
  groupEmail: document.getElementById("group-email"),
  groupEmailProposalsHint: document.getElementById("group-email-proposals-hint"),
  groupCheckIn: document.getElementById("group-check-in"),
  groupCheckInPicker: document.getElementById("group-check-in-picker"),
  groupCheckOut: document.getElementById("group-check-out"),
  groupCheckOutPicker: document.getElementById("group-check-out-picker"),
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
  groupsRows: document.getElementById("groups-rows"),
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
  servicesRows: document.getElementById("services-rows"),
  servicesCount: document.getElementById("services-count"),
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
  reviewsResumeRows: document.getElementById("reviews-resume-rows"),
  reviewsResumeStatus: document.getElementById("reviews-resume-status"),
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
  if (!canApp("communications") && canApp("lost-found")) state.currentView = "lost-found";
  else if (!canApp("communications") && !canApp("lost-found") && canApp("groups")) state.currentView = "groups";
  else if (!canApp("communications") && !canApp("lost-found") && !canApp("groups") && canApp("services")) state.currentView = "services";
  else if (!canApp("communications") && !canApp("lost-found") && !canApp("groups") && !canApp("services") && canApp("reviews")) state.currentView = "reviews";
  else if (!canApp("communications") && state.access.settingsFeatures.length > 0) state.currentView = "settings";
  if (!canSettings("communications") && canSettings("reviews")) state.settingsSection = "reviews";
  else if (!canSettings("communications") && !canSettings("reviews") && canSettings("groups")) state.settingsSection = "groups";
  else if (!canSettings("communications") && !canSettings("reviews") && !canSettings("groups") && canSettings("services")) state.settingsSection = "services";
  else if (!canSettings("communications") && canSettings("admin-users")) state.settingsSection = "admin-users";
  renderLayout();
  renderSettingsSection();
  renderCategoryFilterOptions();
  renderReviewPropertyOptions();
  render();
  await ensureCurrentViewData();
  startAutoRefresh();
}

function bindEvents() {
  els.navCommunications.addEventListener("click", () => setView("communications"));
  els.navLostFound.addEventListener("click", () => setView("lost-found"));
  els.navReviews.addEventListener("click", () => setView("reviews"));
  els.navGroups.addEventListener("click", () => setView("groups"));
  els.navServices.addEventListener("click", () => setView("services"));
  els.openSettings.addEventListener("click", () => setView("settings"));
  els.closeSettings.addEventListener("click", () => setView("communications"));
  els.closeSettingsAdmin.addEventListener("click", () => setView("communications"));
  els.closeSettingsReviews.addEventListener("click", () => setView("reviews"));
  els.closeSettingsGroups.addEventListener("click", () => setView("groups"));
  els.closeSettingsServices.addEventListener("click", () => setView("services"));
  els.settingsMenuCommunications.addEventListener("click", () => setSettingsSection("communications"));
  els.settingsMenuReviews.addEventListener("click", () => setSettingsSection("reviews"));
  els.settingsMenuGroups.addEventListener("click", () => setSettingsSection("groups"));
  els.settingsMenuServices.addEventListener("click", () => setSettingsSection("services"));
  els.settingsMenuAdminUsers.addEventListener("click", () => setSettingsSection("admin-users"));
  els.settingsReviewsImportTab.addEventListener("click", () => setReviewSettingsScreen("import"));
  els.settingsReviewsConfigTab.addEventListener("click", () => setReviewSettingsScreen("config"));
  els.rows.addEventListener("click", onRowAction);
  els.rows.addEventListener("input", onRowDraftInput);
  els.rows.addEventListener("keydown", onRowKeydown);
  els.rows.addEventListener("change", onRowStatusToggle);
  els.lostFoundRows.addEventListener("click", onLostFoundAction);
  els.lostFoundRows.addEventListener("input", onLostFoundDraftInput);
  els.lostFoundRows.addEventListener("keydown", onLostFoundKeydown);
  els.lostFoundRows.addEventListener("change", onLostFoundStatusToggle);
  els.tableHead.addEventListener("click", onSortToggle);
  els.resetSort.addEventListener("click", () => {
    resetSortDefault();
    render();
    showToast("Default sort applied: Date/Time newest first.", "info");
  });
  [els.search, els.showActive, els.statusFilter, els.categoryFilter, els.fromDate, els.toDate].forEach((el) =>
    el.addEventListener("input", render)
  );
  els.showActive.addEventListener("change", render);
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
  els.groupCloseModal.addEventListener("click", closeGroupModal);
  els.groupEditorModal.addEventListener("click", (event) => {
    if (event.target === els.groupEditorModal) closeGroupModal();
  });
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
  [els.groupCheckInPicker, els.groupCheckOutPicker].forEach((el) =>
    el.addEventListener("input", onGroupDatePickerInput)
  );
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
  els.groupsRows.addEventListener("click", onGroupRowClick);
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
  els.serviceTabDetails.addEventListener("click", () => setServiceEditorTab("details"));
  els.serviceTabConfirmation.addEventListener("click", () => setServiceEditorTab("confirmation"));
  els.serviceConfirmationLanguage.addEventListener("input", () => {
    state.serviceConfirmationLanguage = normalizeProposalLanguage(els.serviceConfirmationLanguage.value);
    state.serviceDraft.language = state.serviceConfirmationLanguage;
    renderServiceConfirmationPreview();
  });
  els.servicesRows.addEventListener("click", onServiceRowClick);
  els.serviceCloseModal.addEventListener("click", closeServiceModal);
  els.serviceEditorModal.addEventListener("click", (event) => {
    if (event.target === els.serviceEditorModal) closeServiceModal();
  });
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

async function setView(view) {
  if (view === "settings" && !state.access.settingsFeatures.length) return showToast("No settings access.", "error");
  if (view === "lost-found" && !canApp("lost-found")) return showToast("No Lost&Found access.", "error");
  if (view === "reviews" && !canApp("reviews")) return showToast("No reviews access.", "error");
  if (view === "groups" && !canApp("groups")) return showToast("No groups access.", "error");
  if (view === "services" && !canApp("services")) return showToast("No services access.", "error");
  state.currentView = view;
  if (view === "settings") {
    if (canSettings("communications")) state.settingsSection = "communications";
    else if (canSettings("reviews")) state.settingsSection = "reviews";
    else if (canSettings("groups")) state.settingsSection = "groups";
    else if (canSettings("services")) state.settingsSection = "services";
    else if (canSettings("admin-users")) state.settingsSection = "admin-users";
  }
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
    }
  } finally {
    state.autoRefreshRunning = false;
  }
}

function shouldSkipAutoRefresh() {
  if (state.currentView === "settings") return true;
  if (state.currentView === "communications" && (state.editingId || hasCommunicationDraft())) return true;
  if (state.currentView === "lost-found" && (state.lostFoundEditingId || hasLostFoundDraft())) return true;
  if (state.currentView === "groups" && els.groupEditorModal && !els.groupEditorModal.hidden) return true;
  if (state.currentView === "services" && els.serviceEditorModal && !els.serviceEditorModal.hidden) return true;
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
  if (canSettings("communications") && !state.communicationsSettingsLoaded) {
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
  renderServices();
  renderServiceSettings();
}

async function ensureSettingsSectionData() {
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
  if (state.settingsSection === "admin-users") await ensureAdminUsersData();
}

function renderLayout() {
  const comm = state.currentView === "communications";
  const lostFound = state.currentView === "lost-found";
  const reviews = state.currentView === "reviews";
  const groups = state.currentView === "groups";
  const services = state.currentView === "services";
  const settingsMode = state.currentView === "settings";
  const canComm = canApp("communications");
  const canLostFound = canApp("lost-found");
  const canReviews = canApp("reviews");
  const canGroups = canApp("groups");
  const canServices = canApp("services");

  els.appShell.classList.toggle("settings-mode", settingsMode);
  els.navCommunications.classList.toggle("active", comm);
  els.navLostFound.classList.toggle("active", lostFound);
  els.navReviews.classList.toggle("active", reviews);
  els.navGroups.classList.toggle("active", groups);
  els.navServices.classList.toggle("active", services);
  els.navCommunications.hidden = !canComm;
  els.navLostFound.hidden = !canLostFound;
  els.navReviews.hidden = !canReviews;
  els.navGroups.hidden = !canGroups;
  els.navServices.hidden = !canServices;
  els.openSettings.hidden = !state.access.settingsFeatures.length;
  els.leftNav.hidden = settingsMode;
  els.topbar.hidden = false;
  els.viewCommunications.hidden = !comm;
  els.viewLostFound.hidden = !lostFound;
  els.viewReviews.hidden = !reviews;
  els.viewGroups.hidden = !groups;
  els.viewServices.hidden = !services;
  els.viewSettings.hidden = !settingsMode;
  els.settingsMenuCommunications.hidden = !canSettings("communications");
  els.settingsMenuReviews.hidden = !canSettings("reviews");
  els.settingsMenuGroups.hidden = !canSettings("groups");
  els.settingsMenuServices.hidden = !canSettings("services");
  els.settingsMenuAdminUsers.hidden = !canSettings("admin-users");
  els.settingsMenuCommunications.classList.toggle("active", state.settingsSection === "communications");
  els.settingsMenuReviews.classList.toggle("active", state.settingsSection === "reviews");
  els.settingsMenuGroups.classList.toggle("active", state.settingsSection === "groups");
  els.settingsMenuServices.classList.toggle("active", state.settingsSection === "services");
  els.settingsMenuAdminUsers.classList.toggle("active", state.settingsSection === "admin-users");
}

async function setSettingsSection(section) {
  if (section === "communications" && !canSettings("communications")) return;
  if (section === "reviews" && !canSettings("reviews")) return;
  if (section === "groups" && !canSettings("groups")) return;
  if (section === "services" && !canSettings("services")) return;
  if (section === "admin-users" && !canSettings("admin-users")) return;
  state.settingsSection = section === "admin-users"
    ? "admin-users"
    : section === "services"
      ? "services"
      : section === "groups"
        ? "groups"
        : section === "reviews"
          ? "reviews"
          : "communications";
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
  const isComm = state.settingsSection === "communications" && canSettings("communications");
  const isReviews = state.settingsSection === "reviews" && canSettings("reviews");
  const isGroups = state.settingsSection === "groups" && canSettings("groups");
  const isServices = state.settingsSection === "services" && canSettings("services");
  const isAdmin = state.settingsSection === "admin-users" && canSettings("admin-users");
  els.settingsViewCommunications.hidden = !isComm;
  els.settingsViewReviews.hidden = !isReviews;
  els.settingsViewGroups.hidden = !isGroups;
  els.settingsViewServices.hidden = !isServices;
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
  } catch (e) {
    state.settings = clone(DEFAULT_SETTINGS);
    setSettingsStatus(`Using defaults (${e.message}).`);
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
    showToast("Settings saved.", "success");
  } catch (e) {
    setSettingsStatus(`Save failed: ${e.message}`);
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
      <td class="row-actions">
        <button type="button" class="ghost" data-action="save-user-profile" data-id="${escape(user.id)}">Save</button>
        <button type="button" class="ghost" data-action="reset-user-password" data-id="${escape(user.id)}">Reset Password</button>
      </td>`;
    const select = tr.querySelector("select");
    if (select) select.value = clean(user.profileId);
    els.adminUsersBody.appendChild(tr);
  });
}

function renderProfiles() {
  els.profilesBody.innerHTML = "";
  if (state.profiles.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td colspan="12" class="empty">No profiles yet.</td>';
    els.profilesBody.appendChild(tr);
    return;
  }
  state.profiles.forEach((profile) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input data-profile-name="${escape(profile.id)}" value="${escape(profile.name)}" /></td>
      <td><input type="checkbox" data-profile-app-communications="${escape(profile.id)}" ${profile.appFeatures.includes("communications") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-app-lost-found="${escape(profile.id)}" ${profile.appFeatures.includes("lost-found") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-app-reviews="${escape(profile.id)}" ${profile.appFeatures.includes("reviews") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-app-groups="${escape(profile.id)}" ${profile.appFeatures.includes("groups") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-app-services="${escape(profile.id)}" ${profile.appFeatures.includes("services") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-settings-communications="${escape(profile.id)}" ${profile.settingsFeatures.includes("communications") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-settings-reviews="${escape(profile.id)}" ${profile.settingsFeatures.includes("reviews") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-settings-groups="${escape(profile.id)}" ${profile.settingsFeatures.includes("groups") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-settings-services="${escape(profile.id)}" ${profile.settingsFeatures.includes("services") ? "checked" : ""} /></td>
      <td><input type="checkbox" data-profile-settings-admin-users="${escape(profile.id)}" ${profile.settingsFeatures.includes("admin-users") ? "checked" : ""} /></td>
      <td class="row-actions">
        <button type="button" class="ghost" data-action="save-profile" data-id="${escape(profile.id)}">Save</button>
        <button type="button" class="danger" data-action="delete-profile" data-id="${escape(profile.id)}">Delete</button>
      </td>`;
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
  if (els.profilesBody.querySelector(`[data-profile-app-lost-found="${id}"]`)?.checked) appFeatures.push("lost-found");
  if (els.profilesBody.querySelector(`[data-profile-app-reviews="${id}"]`)?.checked) appFeatures.push("reviews");
  if (els.profilesBody.querySelector(`[data-profile-app-groups="${id}"]`)?.checked) appFeatures.push("groups");
  if (els.profilesBody.querySelector(`[data-profile-app-services="${id}"]`)?.checked) appFeatures.push("services");
  const settingsFeatures = [];
  if (els.profilesBody.querySelector(`[data-profile-settings-communications="${id}"]`)?.checked) settingsFeatures.push("communications");
  if (els.profilesBody.querySelector(`[data-profile-settings-reviews="${id}"]`)?.checked) settingsFeatures.push("reviews");
  if (els.profilesBody.querySelector(`[data-profile-settings-groups="${id}"]`)?.checked) settingsFeatures.push("groups");
  if (els.profilesBody.querySelector(`[data-profile-settings-services="${id}"]`)?.checked) settingsFeatures.push("services");
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

function renderGroups() {
  if (!els.groupsRows || !canApp("groups")) return;
  if (!els.groupEditorModal.hidden) renderGroupDraft();
  const rows = getFilteredGroups();
  updateGroupSortIndicators();
  els.groupsCount.textContent = `${rows.length} proposal${rows.length === 1 ? "" : "s"}`;
  els.groupsRows.innerHTML = "";
  if (rows.length === 0) {
    els.groupsRows.innerHTML = '<tr><td colspan="9" class="empty">No group proposals found.</td></tr>';
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
  });
}

function getFilteredGroups() {
  const today = formatDate(new Date());
  return state.groups
    .filter((row) => !state.groupsShowActive || clean(row.checkOut) >= today)
    .sort(compareGroupRows);
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
    roomTypes: groupRoomTypeSummary(row.roomItems),
    rooms: groupRoomSelectionSummary(row.roomItems),
    optionDate: row.optionDate || "-",
    reservationNumber: row.reservationNumber || "-",
    observation: row.observation || "",
  }));
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
  const tableRows = rows.map((row) => `<tr class="${row.status === "Accepted" ? "accepted" : row.status === "Refused" ? "refused" : ""}">
    <td>${escape(row.created)}</td>
    <td>${escape(row.name)}</td>
    <td>${escape(row.checkIn)} - ${escape(row.checkOut)}<br><small>${escape(String(row.nights))} nights</small></td>
    <td>${escape(String(row.guests))}</td>
    <td>${escape(row.language)}</td>
    <td><strong>${escape(row.total)}</strong><br><small>${escape(row.deposit)}</small></td>
    <td>${escape(row.roomTypes).replaceAll(", ", "<br>")}</td>
    <td>${escape(row.rooms || "-")}</td>
    <td>${escape(row.optionDate)}</td>
    <td>${escape(row.reservationNumber)}</td>
  </tr>`).join("");
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
          body > p:not(.summary) { display: none; }
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
        <p class="summary">Exported ${escape(date)} &middot; ${escape(String(rows.length))} proposal${rows.length === 1 ? "" : "s"} &middot; ${state.groupsShowActive ? "Active only" : "All proposals"}</p>
        <p>Exported ${escape(date)} · ${escape(String(rows.length))} proposal${rows.length === 1 ? "" : "s"} · ${state.groupsShowActive ? "Active only" : "All proposals"}</p>
        <table>
          <thead>
            <tr><th>Created</th><th>Name</th><th>Dates</th><th>Guests</th><th>Language</th><th>Total</th><th>Room Types</th><th>Rooms</th><th>Option</th><th>Reservation</th></tr>
          </thead>
          <tbody>${tableRows || '<tr><td colspan="10">No group proposals found.</td></tr>'}</tbody>
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
  els.groupCheckIn.value = formatGroupDateInput(draft.checkIn);
  els.groupCheckInPicker.value = draft.checkIn || "";
  els.groupCheckOut.value = formatGroupDateInput(draft.checkOut);
  els.groupCheckOutPicker.value = draft.checkOut || "";
  els.groupGuests.value = draft.guests;
  els.groupOptionDate.value = draft.optionDate;
  els.groupStatusField.value = draft.status;
  renderGroupStatusColor();
  els.groupObservation.value = draft.observation;
  renderGroupRoomItems();
  renderGroupTotals();
  renderGroupAuditHistory();
  els.groupDelete.hidden = !draft.id;
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
  state.groupDraft.guests = normalizeGroupGuests(els.groupGuests.value);
  state.groupDraft.optionDate = clean(els.groupOptionDate.value);
  state.groupDraft.status = normalizeGroupStatus(els.groupStatusField.value);
  state.groupDraft.observation = clean(els.groupObservation.value);
  syncGroupDatePickers();
  renderGroupStatusColor();
  renderGroupTotals();
}

function onGroupDatePickerInput(event) {
  if (event.target === els.groupCheckInPicker) {
    state.groupDraft.checkIn = clean(els.groupCheckInPicker.value);
    els.groupCheckIn.value = formatGroupDateInput(state.groupDraft.checkIn);
  }
  if (event.target === els.groupCheckOutPicker) {
    state.groupDraft.checkOut = clean(els.groupCheckOutPicker.value);
    els.groupCheckOut.value = formatGroupDateInput(state.groupDraft.checkOut);
  }
  renderGroupTotals();
}

function syncGroupDatePickers() {
  els.groupCheckInPicker.value = state.groupDraft.checkIn || "";
  els.groupCheckOutPicker.value = state.groupDraft.checkOut || "";
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
  const row = event.target.closest("tr[data-group-id]");
  if (!row) return;
  const group = state.groups.find((item) => item.id === clean(row.dataset.groupId));
  if (!group) return;
  await refreshGroupSettingsForEditor();
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
  return `${raw.slice(8, 10)}/${raw.slice(5, 7)}/${raw.slice(0, 4)}`;
}

function formatGroupDateDisplay(value) {
  const raw = clean(value);
  return raw ? formatGroupDateInput(raw) : "-";
}

function parseGroupDateInput(value) {
  const raw = clean(value);
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
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

async function loadLostFound({ silent = false } = {}) {
  try {
    const result = await api("/api/lost-found");
    state.lostFound = (result.rows || []).map((row) => ({
      id: row.id,
      number: Number(row.item_number) || 0,
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
    t.closest("tr")?.style && (t.closest("tr").style.backgroundColor = "#ffffff");
  }
  if (scope === "edit" && state.editingId && t.dataset.id === state.editingId) {
    state.editDraft[field] = value;
    t.closest("tr")?.style && (t.closest("tr").style.backgroundColor = rowBackgroundColor(state.editDraft.status, state.editDraft.category));
  }
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
  const entry = state.entries.find((x) => x.id === id);
  if (!entry) return;
  if (action === "edit") {
    state.editingId = id;
    state.editDraft = { person: entry.person, status: entry.status, category: entry.category, message: entry.message };
    return render();
  }
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
  const record = state.lostFound.find((item) => item.id === id);
  if (!record) return;
  if (action === "edit") {
    state.lostFoundEditingId = id;
    state.lostFoundEditDraft = {
      whoFound: record.whoFound,
      whoRecorded: record.whoRecorded,
      location: record.location,
      objectDescription: record.objectDescription,
      notes: record.notes,
      stored: record.stored,
      status: record.status,
    };
    renderLostFound();
    return;
  }
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
    const row = target.closest("tr");
    if (row) row.style.backgroundColor = "#ffffff";
  }
  if (scope === "edit" && state.lostFoundEditingId && clean(target.dataset.id) === state.lostFoundEditingId) {
    state.lostFoundEditDraft[field] = value;
    const row = target.closest("tr");
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
  const configs = Array.isArray(settings?.serviceConfigs || settings?.service_configs)
    ? settings.serviceConfigs || settings.service_configs
    : [];
  output.serviceConfigs = configs.map(sanitizeServiceConfigClient).filter((item) => item.serviceType);
  if (!output.serviceConfigs.length) output.serviceConfigs = clone(DEFAULT_SERVICE_SETTINGS.serviceConfigs);
  return output;
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
    leg: "arrival",
    pickupLocation: draft.pickupLocation,
    dropoffLocation: draft.dropoffLocation,
  };
  const returnOptions = {
    flightNumber: draft.returnFlight,
    date: draft.returnDate,
    leg: "departure",
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
  return `${raw.slice(0, 4)}/${raw.slice(5, 7)}/${raw.slice(8, 10)}`;
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
  const pickup = clean(row.pickupLocation).toLowerCase();
  const dropoff = clean(row.dropoffLocation).toLowerCase();
  const today = formatDate(new Date());
  const airportPattern = /(airport|aeroporto|humberto delgado|portela|lisbon airport|aeroporto de lisboa|terminal 1|terminal 2|lis)/i;
  if (!clean(row.flightNumber) || clean(row.date) !== today) return false;
  return row.legType === "return" ? airportPattern.test(dropoff) : airportPattern.test(pickup);
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
  const statusText = serviceFlightStatusText(row);
  if (!statusText) return escape(flight);
  return `${escape(flight)}<small class="service-flight-status-note">(${escape(statusText)})</small>`;
}

async function fetchServiceFlightStatus(row) {
  const key = serviceFlightStatusKey(row);
  if (!serviceAirportArrivalCandidate(row)) {
    delete state.serviceFlightStatuses.cache[key];
    return;
  }
  if (state.serviceFlightStatuses.cache[key]?.loaded || state.serviceFlightStatuses.cache[key]?.pending) return;
  state.serviceFlightStatuses.cache[key] = { text: "checking...", loaded: false, pending: true };
  renderServices();
  try {
    const leg = row.legType === "return" ? "departure" : "arrival";
    const result = await api(`/api/aviationstack-flight?flight=${encodeURIComponent(normalizeFlightCode(row.flightNumber))}&date=${encodeURIComponent(clean(row.date))}&leg=${encodeURIComponent(leg)}&pickup=${encodeURIComponent(clean(row.pickupLocation))}&dropoff=${encodeURIComponent(clean(row.dropoffLocation))}`);
    const timeText = formatPredictionTime(result?.predictedTime, result?.timeZone);
    const label = row.legType === "return" ? "ETD" : "ETA";
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

function renderServices() {
  if (!els.servicesRows || !canApp("services")) return;
  els.servicesShowActive.checked = !!state.serviceFilters.showActive;
  els.servicesFilterCreatedFrom.value = clean(state.serviceFilters.createdFrom);
  els.servicesFilterCreatedTo.value = clean(state.serviceFilters.createdTo);
  els.servicesFilterDateFrom.value = clean(state.serviceFilters.dateFrom);
  els.servicesFilterDateTo.value = clean(state.serviceFilters.dateTo);
  els.servicesFilterName.value = clean(state.serviceFilters.name);
  const rows = getFilteredServices();
  els.servicesCount.textContent = `${rows.length} service line${rows.length === 1 ? "" : "s"}`;
  els.servicesRows.innerHTML = "";
  if (!rows.length) {
    els.servicesRows.innerHTML = '<tr><td colspan="11" class="empty">No services found.</td></tr>';
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
      <td><span class="service-status-pill ${serviceStatusTone(row.status)}">${escape(row.status || "-")}</span></td>
      <td>${renderServiceDateCell(row.date)}</td>
      <td>${escape(clean(row.time) || "-")}</td>
      <td>${escape(String(row.pax || 0))}</td>
      <td>${escape(row.pickupLocation || "-")}</td>
      <td>${renderServiceFlightCell(row)}</td>
      <td>${escape(row.dropoffLocation || "-")}</td>
      <td>${escape(formatMoney(row.price))}</td>`;
    els.servicesRows.appendChild(tr);
  });
  refreshVisibleServiceStatuses();
}

function renderServiceSettings() {
  if (!els.servicesConfigsBody) return;
  const configs = serviceConfigs();
  if (els.servicesAutomaticEmailRecipients) {
    els.servicesAutomaticEmailRecipients.value = (state.serviceSettings?.automaticEmailRecipients || []).join("\n");
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
}

function onServiceRowClick(event) {
  const row = event.target.closest("tr[data-service-id]");
  if (!row) return;
  const service = state.services.find((item) => item.id === clean(row.dataset.serviceId));
  if (!service) return;
  state.serviceSelectedId = service.id;
  state.serviceDraft = { ...clone(service), language: normalizeServiceConfirmationLanguage(service.language), priceManual: false };
  openServiceModal();
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
  if (!clean(draft.date)) return setServicesStatus("Date is required.");
  if (!clean(draft.time)) return setServicesStatus("Time is required.");
  if (clean(draft.customerPhone) && !isValidInternationalPhone(draft.customerPhone)) {
    return setServicesStatus("Customer Phone must include country code, for example +351 912 345 678.");
  }
  if (draft.hasReturn && clean(draft.returnDate) && clean(draft.date) && clean(draft.returnDate) < clean(draft.date)) {
    return setServicesStatus("Return date cannot be before the main service date.");
  }
  const previous = draft.id ? state.services.find((item) => item.id === draft.id) : null;
  const payload = {
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

function render() {
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
  if (!canApp("communications")) {
    els.count.textContent = "0 records";
    els.rows.innerHTML = '<tr><td colspan="7" class="empty">Your profile has no access to Communications.</td></tr>';
    return;
  }
  const rows = getFilteredEntries();
  els.count.textContent = `${rows.length} record${rows.length === 1 ? "" : "s"}`;
  els.rows.innerHTML = "";
  els.rows.appendChild(buildInlineRow());
  if (rows.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="7" class="empty">No communications found.</td>`;
    els.rows.appendChild(tr);
    updateSortIndicators();
    syncStickyRows();
    return;
  }
  rows.forEach((entry) => els.rows.appendChild(state.editingId === entry.id ? buildEditableRow(entry) : buildReadOnlyRow(entry)));
  updateSortIndicators();
  syncStickyRows();
}

function renderLostFound() {
  if (!canApp("lost-found")) {
    els.lostFoundCount.textContent = "0 records";
    els.lostFoundRows.innerHTML = '<tr><td colspan="7" class="empty">Your profile has no access to Lost&Found.</td></tr>';
    return;
  }
  const rows = getFilteredLostFoundRecords();
  els.lostFoundCount.textContent = `${rows.length} record${rows.length === 1 ? "" : "s"}`;
  els.lostFoundRows.innerHTML = "";
  els.lostFoundRows.appendChild(buildLostFoundInlineRow());
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

function buildLostFoundInlineRow() {
  const draft = state.lostFoundDraft || emptyLostFoundDraft();
  const now = new Date();
  const tr = document.createElement("tr");
  tr.className = "inline-editor sticky-new-row";
  tr.innerHTML = `<td class="lost-found-timestamp"><span class="auto-stamp">Auto</span><small>${formatDate(now)} ${formatTime(now)}</small></td>
    <td><input data-field="whoFound" data-scope="new" value="${escape(draft.whoFound)}" placeholder="Found" />
    <input data-field="whoRecorded" data-scope="new" value="${escape(draft.whoRecorded)}" placeholder="Record" /></td>
    <td><input data-field="location" data-scope="new" value="${escape(draft.location)}" placeholder="Where" />
    <select data-field="stored" data-scope="new">${LOST_FOUND_STORED_OPTIONS.map((item) => option(item, draft.stored)).join("")}</select></td>
    <td><input data-field="objectDescription" data-scope="new" value="${escape(draft.objectDescription)}" /></td>
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
    <td class="message">${escape(record.objectDescription)}</td>
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
    <td><input data-field="objectDescription" data-scope="edit" data-id="${escape(record.id)}" value="${escape(draft.objectDescription)}" /></td>
    <td><textarea data-field="notes" data-scope="edit" data-id="${escape(record.id)}" rows="2">${escape(draft.notes)}</textarea></td>
    <td class="lost-found-status-cell"><label class="status-toggle"><input type="checkbox" data-field="status" data-scope="edit" data-id="${escape(record.id)}" ${isClosedStatus(draft.status) ? "checked" : ""} /><span>Closed</span></label></td>
    <td class="row-actions lost-found-actions-compact"><button type="button" data-lost-found-action="save-edit" data-id="${escape(record.id)}">Save</button>
    <button type="button" data-lost-found-action="cancel-edit" data-id="${escape(record.id)}" class="ghost">Cancel</button></td>`;
  tr.style.backgroundColor = lostFoundRowBackground(draft.status);
  return tr;
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
  const input = settings?.communications || {};
  const categories = Array.isArray(input.categories) ? input.categories : [];
  const seen = new Set();
  output.communications.categories = categories
    .map((c) => ({ name: clean(c?.name), color: normalizeHex(c?.color), autoCloseDays: normalizeAutoCloseDays(c?.autoCloseDays) }))
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

function normalizeTimeInput(value) {
  return /^\d{2}:\d{2}$/.test(clean(value)) ? clean(value) : "00:00";
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
  renderReviewDetail(rows);
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
    els.reviewsGoogleMappingsBody.innerHTML = '<tr><td colspan="2" class="empty">Add a property before mapping Google locations.</td></tr>';
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
      <td><select data-google-property-location="${escape(property.id)}">${locationOptions}</select></td>`;
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
  if (rows.length === 0) {
    els.reviewsRows.innerHTML = '<tr><td colspan="6" class="empty">No reviews match the current filters.</td></tr>';
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

function renderReviewDetail(rows) {
  if (!els.reviewsDetail) return;
  const selected = rows.find((row) => clean(row.id) === clean(state.reviewSelectedId)) || rows[0] || null;
  if (!selected) {
    state.reviewSelectedId = "";
    els.reviewsDetail.className = "review-detail empty";
    els.reviewsDetail.textContent = "Select a review to see the full detail.";
    return;
  }
  state.reviewSelectedId = clean(selected.id);
  const meta = [
    ["Date", clean(selected.review_date) || "-"],
    ["Property", clean(selected.properties?.name || reviewPropertyName(selected.property_id) || "-")],
    ["Reviewer", clean(selected.reviewer_name) || "Anonymous"],
    ["Source", reviewSourceLabel(selected.source)],
    ["Score", formatReviewScore(selected.rating_normalized_100, selected.rating_raw, selected.rating_scale)],
  ];
  const reservationValue = clean(selected.source_reservation_id) || (clean(selected.source).toLowerCase() === "hostelworld" ? clean(selected.source_review_id) : "");
  if (reservationValue) meta.push(["Booking / Reservation", reservationValue]);
  else if (clean(selected.source_review_id)) meta.push(["Source Reference", clean(selected.source_review_id)]);
  const detail = [
    `<p class="review-detail-title"><strong>${escape(clean(selected.title) || "(no title)")}</strong></p>`,
    `<div class="review-detail-meta">${meta.map(([label, value]) => `<div class="review-detail-meta-item"><span>${escape(label)}</span><strong>${escape(value)}</strong></div>`).join("")}</div>`,
  ];
  const subscoreHtml = renderReviewSubscores(selected.subscores);
  if (subscoreHtml) detail.push(`<div class="review-detail-section"><strong>Partial scores:</strong>${subscoreHtml}</div>`);
  if (clean(selected.positive_review_text)) detail.push(`<p class="review-detail-section"><strong>Positive:</strong> ${escape(clean(selected.positive_review_text))}</p>`);
  if (clean(selected.negative_review_text)) detail.push(`<p class="review-detail-section"><strong>Negative:</strong> ${escape(clean(selected.negative_review_text))}</p>`);
  if (clean(selected.body)) detail.push(`<p class="review-detail-section"><strong>Full review:</strong><br />${escape(clean(selected.body))}</p>`);
  if (clean(selected.host_reply_text)) detail.push(`<p class="review-detail-section"><strong>Property reply:</strong><br />${escape(clean(selected.host_reply_text))}</p>`);
  els.reviewsDetail.className = "review-detail";
  els.reviewsDetail.innerHTML = detail.join("");
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
  const row = event.target.closest("tr[data-review-id]");
  if (!row) return;
  state.reviewSelectedId = clean(row.dataset.reviewId);
  renderReviews();
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
  try {
    setReviewGoogleStatus("Syncing Google reviews...");
    const result = await api("/api/google-business?action=sync", { method: "POST", body: {} });
    await Promise.all([loadReviews({ useFilters: true }), loadReviewImportRuns()]);
    render();
    const inserted = Number(result.insertedCount || 0);
    const replaced = Number(result.replacedCount || 0);
    const imported = Number(result.importedCount || 0);
    setReviewGoogleStatus(`Google sync complete: ${imported} review${imported === 1 ? "" : "s"} processed, ${inserted} inserted, ${replaced} replaced.`);
    showToast("Google reviews synced.", "success");
  } catch (e) {
    setReviewGoogleStatus(`Google sync failed: ${e.message}`);
    showToast(`Google sync failed: ${e.message}`, "error");
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

function setGroupsStatus(text) {
  if (els.groupsStatus) els.groupsStatus.textContent = text;
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
