const entries = require("./data.json").log.entries;
const xl = require("excel4node");
var base64 = require("base-64");
const url = require("url");
var utf8 = require("utf8");
const qs = require("qs");
const { verify } = require("crypto");
const wb = new xl.Workbook();
const wsDataRaw = wb.addWorksheet("DATA RAW");
const wsLinks = wb.addWorksheet("LINKS");
const picpath = [
  "speech-2",
  "car-2",
  "drone-2",
  "bomb-1",
  "ak-2",
  "truck-1",
  "rescue-2",
  "phone-2",
  "fires-1",
  "arrested-2",
  "speech-11",
  "rally-2",
  "phone-1",
  "explode-2",
  "capture-1",
  "ak-1",
  "money-2",
  "dead-2",
  "civil_airplane-1",
  "arrested-1",
  "capture-2",
  "airplane-2",
  "speech-5",
  "airplane-1",
  "heavy-1",
  "bomb-2",
  "flares-2",
  "speech-1",
  "rocket-1",
  "aa-2",
  "wifi-1",
  "helicopter-1",
  "civil_airplane-2",
  "speech-7",
  "destroy-1",
  "explode-1",
  "ship-1",
  "stop-1",
  "gas-1",
  "fires-2",
  "video-2",
  "truck-2",
  "picture-1",
  "ak-11",
  "police-2",
  "medicine-2",
  "dead-1",
  "elect-1",
  "satellite-2",
  "video-1",
  "phone-7",
  "civil_airplane-7",
  "molotov-2",
  "manpads-2",
  "destroy-2",
  "artillery-1",
  "comp-2",
  "rally-4",
  "food-2",
  "stop-7",
  "natural_resource-2",
  "aa-1",
  "ship-11",
  "drone-1",
  "press-1",
  "supply-2",
  "natural_resource-1",
  "fires-4",
  "press-2",
  "speech-4",
  "thug-1",
  "railway-2",
  "stop-2",
  "money-1",
  "medicine-1",
  "map-1",
  "helicopter-2",
  "dead-11",
  "dead-5",
  "comp-1",
  "railway-1",
  "mobile-1",
  "food-1",
  "ship-2",
  "bus-2",
  "camp-1",
  "rally-1",
  "heavy-2",
  "artillery-2",
  "wifi-2",
  "manpads-1",
  "elect-2",
  "gun-2",
  "speech-9",
  "twitterico-5",
  "floods-2",
  "atgm-2",
  "facebookico-2",
  "mine-1",
  "lightplane-1",
  "twitterico-2",
  "molotov-1",
  "picture-2",
  "hostage-1",
  "hostage-2",
  "map-2",
  "phone-9",
  "phone-11",
  "floods-5",
  "rally-6",
  "gas-2",
  "polution-2",
  "police-9",
  "polution-1",
  "speech-10",
  "ship-7",
  "crane-2",
  "airplane-6",
  "flares-1",
  "car-1",
  "mine-2",
  "floods-1",
  "police-1",
  "bus-5",
  "twitterico-1",
  "rescue-1",
  "nowater-2",
  "gun-1",
  "submarine-1",
  "nowater-1",
  "crane-1",
  "atgm-1",
  "food-11",
  "rocket-2",
  "ship-5",
  "earthquake-5",
  "explode-4",
  "car-5",
  "fires-5",
  "facebookico-1",
  "supply-1",
];
let linkRow = 2;
wsLinks.cell(1, 1).string("Type");
wsLinks.cell(1, 2).string("Date");
wsLinks.cell(1, 3).string("Longitude");
wsLinks.cell(1, 4).string("Latitude");
wsLinks.cell(1, 5).string("Location");
wsLinks.cell(1, 6).string("Link");
wsLinks.cell(1, 7).string("Source");
const slice = entries
  .map((entry) => {
    const result = {};
    try {
      entry.response.content.json = JSON.parse(entry.response.content.text);
      delete entry.response.content.text;
    } catch (err) {
      try {
        const bytes = base64.decode(entry.response.content.text);
        const text = utf8.decode(bytes);
        entry.response.content.json = JSON.parse(text);
        delete entry.response.content.text;
      } catch (err) {
        console.error(err);
      }
    }

    result.query = qs.parse(url.parse(entry.request.url).query);
    result.json = entry.response.content.json;
    result.schema = generateSchema(result.json);
    if (result.query.act === "pts") {
      result.stats = {};
      result.markerData = [];
      result.json.venues.forEach((venue) => {
        if (!result.stats[venue.picpath]) result.stats[venue.picpath] = 0;
        result.stats[venue.picpath]++;
        result.markerData.push({
          longitude: venue.lng,
          latitude: venue.lat,
          pic: venue.picpath,
        });
        wsLinks.cell(linkRow, 1).string(venue.picpath);
        wsLinks.cell(linkRow, 2).date(new Date(venue.timestamp * 1000));
        wsLinks.cell(linkRow, 3).string(venue.lng);
        wsLinks.cell(linkRow, 4).string(venue.lat);
        wsLinks.cell(linkRow, 5).string(String(venue.location));
        wsLinks.cell(linkRow, 6).link(String(venue.link));
        wsLinks.cell(linkRow, 7).link(String(venue.source));

        linkRow++;
      });
      result.date = new Date(parseInt(result.query.time) * 1000);
    }

    delete result.json;
    return result;
  })
  .filter((result) => result.query.act === "pts")
  .map((result) => result)
  .flat();

function generateSchema(obj) {
  let result = {};
  for (const [key, value] of Object.entries(obj)) {
    const valueType = typeof value;
    if (
      [null, undefined].includes(value) ||
      ["string", "number", "bigint", "boolean", "symbol", "undefined"].includes(
        valueType
      )
    ) {
      result[key] = value === "null" ? "null" : valueType;
    } else if (valueType === "function") {
      result[key] = `function ${value.name}`;
    } else {
      if (Array.isArray(value)) {
        if (value.length) {
          result[key] = [generateSchema(value[0])];
        } else {
          result[key] = [];
        }
      } else {
        result[key] = generateSchema(value);
      }
    }
  }

  return result;
}
// console.log(JSON.stringify(slice, null, 2));
for (let i = 0; i < slice.length; i++) {
  const result = slice[i];
  wsDataRaw.cell(1, i + 2).date(result.date);
  picpath.forEach((picpath, index) => {
    wsDataRaw.cell(index + 2, i + 2).number(result.stats[picpath] ?? 0);
  });
}
picpath.forEach((picpath, index) => {
  wsDataRaw.cell(index + 2, 1).string(picpath);
});

wb.write(Math.random() + "result.xlsx");
