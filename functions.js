// Koppelt de JS-functie aan de functienaam die Excel gebruikt (moet matchen met functions.json)
CustomFunctions.associate("VANHIERCRM", VANHIERCRM);

/**
 * Haalt een waarde op uit VanhierCRM
 * @param {string} kolomnaam Naam van de kolom
 * @param {string} klantnummer Klantnummer
 * @returns {Promise<string>} Resultaat uit CRM
 */
async function VANHIERCRM(kolomnaam, klantnummer) {
  // TODO: vervang door je eigen Function App naam en function key
  const baseUrl = "https://vanhierinternerapportagepremium.azurewebsites.net/api/klanten";
  const code = "G0Y989WRLvuqJU_aoczo3_RWmxUDyJUCG57J3ewEwbXJAzFu31k47Q==";

  const url =
    `${baseUrl}?code=${code}` +
    `&kolom=${encodeURIComponent(kolomnaam)}` +
    `&klantnummer=${encodeURIComponent(klantnummer)}`;

  const response = await fetch(url);

  if (!response.ok) {
    throw new Error(`CRM fout: ${response.status}`);
  }

  const tekst = await response.text();
  return tekst.trim();
}
