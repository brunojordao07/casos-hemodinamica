// Configuração do app. O CLIENT_ID não é segredo: é o "Application (client) ID" do registro no Azure.
window.APP_CONFIG = {
  CLIENT_ID: "d3a0a6bf-da46-4621-a68c-66446de3237b",
  // Pasta (a partir da raiz do OneDrive) onde ficam casos.json, a planilha e a pasta Imagens.
  ROOT_FOLDER: "Casos Hemodinâmica",
  // Encerrar a sessão após este tempo sem uso (minutos).
  IDLE_MINUTES: 30
};
