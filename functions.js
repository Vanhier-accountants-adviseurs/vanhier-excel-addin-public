// Koppelt de JS-functie aan de functienaam die Excel gebruikt (moet matchen met functions.json)
CustomFunctions.associate("VANHIERCRM", VANHIERCRM);

/**
 * Haalt een waarde op uit VanhierCRM
 * @param {string} kolomnaam Naam van de kolom
 * @param {string} klantnummer Klantnummer
 * @returns {Promise<string>} Resultaat uit CRM
 */
async function VANHIERCRM(kolomnaam, klantnummer) {
  // TODO: vervang door je eigen Function App naam
  const baseUrl = "https://vanhierinternerapportagepremium.azurewebsites.net";

  let token;
  try {
    token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false });
  } catch (e) {
    throw new Error("Niet ingelogd / geen SSO-token: " + e.message);
  }

  const url =
    `${baseUrl}/api/klanten?kolom=${encodeURIComponent(kolomnaam)}` +
    `&klantnummer=${encodeURIComponent(klantnummer)}`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (!response.ok) {
    throw new Error(`CRM fout: ${response.status}`);
  }

  const tekst = await response.text();
  return tekst.trim();
}
