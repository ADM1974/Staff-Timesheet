// Returns active site names for the dropdown (signed-in staff only).
const { getAppToken, validateStaffToken, getActiveSites } = require('./staff');

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
  const user = await validateStaffToken(event.headers && (event.headers.authorization || event.headers.Authorization));
  if (!user) return { statusCode: 401, headers, body: '{"sites":[]}' };
  try {
    const token = await getAppToken();
    return { statusCode: 200, headers, body: JSON.stringify({ sites: await getActiveSites(token) }) };
  } catch (err) {
    console.error(err);
    return { statusCode: 502, headers, body: JSON.stringify({ sites: [], error: 'server' }) };
  }
};
