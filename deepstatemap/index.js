const AreaUtils = {
  RADIUS: 6378137,
  geometry: function geometry(feature) {
    let area = 0,
      i;
    switch (feature.type) {
      case "Polygon":
        return AreaUtils.polygonArea(feature.coordinates);
      case "MultiPolygon":
        for (i = 0; i < feature.coordinates.length; i++) {
          area += AreaUtils.polygonArea(feature.coordinates[i]);
        }
        return area;
      case "Point":
      case "MultiPoint":
      case "LineString":
      case "MultiLineString":
        return 0;
      case "GeometryCollection":
        for (i = 0; i < feature.geometries.length; i++) {
          area += AreaUtils.geometry(feature.geometries[i]);
        }
        return area;
    }
  },
  polygonArea: function polygonArea(coords) {
    var area = 0;

    if (coords && coords.length > 0) {
      area += Math.abs(AreaUtils.ringArea(coords[0]));

      for (var i = 1; i < coords.length; i++) {
        area -= Math.abs(AreaUtils.ringArea(coords[i]));
      }
    }

    return area;
  },
  ringArea: function ringArea(coords) {
    var p1,
      p2,
      p3,
      lowerIndex,
      middleIndex,
      upperIndex,
      i,
      area = 0,
      coordsLength = coords.length;

    if (coordsLength > 2) {
      for (i = 0; i < coordsLength; i++) {
        if (i === coordsLength - 2) {
          // i = N-2
          lowerIndex = coordsLength - 2;
          middleIndex = coordsLength - 1;
          upperIndex = 0;
        } else if (i === coordsLength - 1) {
          // i = N-1
          lowerIndex = coordsLength - 1;
          middleIndex = 0;
          upperIndex = 1;
        } else {
          // i = 0 to N-3
          lowerIndex = i;
          middleIndex = i + 1;
          upperIndex = i + 2;
        }

        p1 = coords[lowerIndex];
        p2 = coords[middleIndex];
        p3 = coords[upperIndex];
        area +=
          (AreaUtils.rad(p3[0]) - AreaUtils.rad(p1[0])) *
          Math.sin(AreaUtils.rad(p2[1]));
      }

      area = (area * AreaUtils.RADIUS * AreaUtils.RADIUS) / 2;
    }

    return area;
  },
  rad: function rad(_) {
    return (_ * Math.PI) / 180;
  },
};

const entries = require("./data_2.json").log.entries;
const xl = require("excel4node");
const wb = new xl.Workbook();
const ws = wb.addWorksheet("DATA RAW");

const colorMap = {
  "#a52714": "Оккупировано рф",
  "#0288d1": "Освобождено Украиной",
  "#bdbdbd": "Идут активные бои/Под вопросом",
  "#880e4f": "Приднестровье/ОРДЛО",
  "#0f9d58": "Освобождено Украиной",
  "#757575": "Идут активные бои/Под вопросом",
  "#ff5252": "Приднестровье/ОРДЛО",
  "#0097a7": "Освобождено Украиной",
  "#000000": "Приднестровье/ОРДЛО",
  "#bcaaa4": "Идут активные бои/Под вопросом",
};
const columns = [...new Set(Object.values(colorMap))];
const results = entries
  .filter((entry) => entry.request.url.includes("/geojson"))
  .map((entry) => {
    const result = {};
    result.date = new Date(+entry.request.url.match(/\d+/)[0] * 1000);
    result.content = JSON.parse(entry.response.content.text)
      .features.filter((e) => e.properties.fill)
      .map((e) => ({
        ...e,
        area: AreaUtils.geometry(e.geometry) / 1000000,
        fill: e.properties.fill,
      }));
    result.colums = Object.fromEntries(columns.map((name) => [name, 0]));
    result.content.forEach((e) => {
      e.type =
        e.fill === "#000000" && e.properties.name === "Під питанням"
          ? "Идут активные бои/Под вопросом"
          : colorMap[e.fill];
      result.colums[e.type] += e.area;
    });
    return result;
  });
ws.cell(1, 1).string("Дата");

columns.forEach((name, index) => ws.cell(1, index + 2).string(name));

for (let i = 0; i < results.length; i++) {
  const row = i + 2;
  const result = results[i];
  ws.cell(row, 1).date(result.date);
  columns.forEach((name, index) =>
    ws.cell(row, index + 2).number(result.colums[name])
  );
}

wb.write(Math.random() + "result.xlsx");
