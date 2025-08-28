"use strict";

// Leaflet & Esri Leaflet
import * as L from "leaflet";
import "leaflet/dist/leaflet.css";
import * as LEsri from "esri-leaflet";
type EsriFeatureLayer = L.esri.FeatureLayer;

// Grouped Layer Control
import "leaflet-groupedlayercontrol";
import "./../style/leaflet.groupedlayercontrol.css";

// Power BI
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

// County GeoJSON + FIPS lookup
import counties from "./data/nc_counties";
import fipsToCounty from "./data/fipsToCounty";

import "./../style/visual.less";

export class Visual implements IVisual {
  private target!: HTMLElement;
  private host!: IVisualHost;
  private mapContainer!: HTMLElement;
  private map!: L.Map;

  private countiesLayer!: L.GeoJSON;
  private hccsLayer?: EsriFeatureLayer;
  private hccsStyle!: (feature: GeoJSON.Feature) => L.PathOptions;

  private currentMode: "LEL" | "COUNTY" = "LEL";
  private lastWhere = "";
  private lastLelKey = "";
  private lastCountyKey = "";
  private lastTroopKey = "";
  private lastLelAltWhere: string | null = null;

  private layerControl?: L.Control.Layers;

  private lelRegionsLayer?: L.esri.FeatureLayer;
  private mpoLayer?: L.esri.FeatureLayer;
  private troopLayer?: L.esri.FeatureLayer;

  // Work zone layers
  private truckClosureLayer?: L.esri.FeatureLayer;
  private constructionLayer?: L.esri.FeatureLayer;
  private nightConstructionLayer?: L.esri.FeatureLayer;
  private maintenanceLayer?: L.esri.FeatureLayer;
  private nightMaintenanceLayer?: L.esri.FeatureLayer;
  private emergencyLayer?: L.esri.FeatureLayer;
  private obstructionLayer?: L.esri.FeatureLayer;
  private weatherLayer?: L.esri.FeatureLayer;
  private specialLayer?: L.esri.FeatureLayer;
  private otherLayer?: L.esri.FeatureLayer;

  // Basemaps
  private darkMap!: L.TileLayer;
  private lightMap!: L.TileLayer;

  // Legend
  private legendControl?: L.Control;

  // Update coalescing
  private updateToken = 0;
  private pendingFitTimer: number | null = null;

  private autoRefreshTimer: number | null = null;

  constructor(options: VisualConstructorOptions | undefined) {
    this.target = options?.element!;
    this.host = options?.host!;
    this.createMapContainer();
    this.initMap();

    // -------- County overlay --------
    this.countiesLayer = L.geoJSON(counties as any, {
      pane: "counties",
      style: { fill: true, fillOpacity: 0.0, fillColor: "#000", weight: 1, color: "gray", dashArray: "3" },
      onEachFeature: (feature, layer) => {
        const fips = (feature.properties as any)?.FIPS;
        if (fips != null) {
          const name = fipsToCounty[Number(fips)] ?? "Unknown";
          (layer as L.Path).bindTooltip(`${name} County`, { sticky: true });
          layer.on("mouseover", () => (layer as L.Path).setStyle({ weight: 3, color: "yellow" }));
          layer.on("mouseout", () => (layer as L.Path).setStyle({ weight: 1, color: "gray" }));
        }
      }
    }).addTo(this.map);

    // -------- Work zone style / tooltip --------
    const workzoneStyle: L.PathOptions = { interactive: true, weight: 10, color: "blue", opacity: 0.3 };
    const workzoneEachFeature = (feature: GeoJSON.Feature, layer: L.Layer) => {
      const p = (feature.properties as any) || {};
      const sev = p.Severity === 1 ? "Low" : p.Severity === 2 ? "Medium" : p.Severity === 3 ? "High" : "Not Provided";
      const dirMap: any = { N: "North", S: "South", E: "East", W: "West", A: "All", O: "Outer" };
      const direction = dirMap[p.Direction] || "Not Provided";
      const label = `<strong>Type:</strong> ${p.IncidentType || "Not Provided"}<br/>
                     <strong>Impact Level:</strong> ${sev}<br/>
                     <strong>Condition:</strong> ${p.Condition || "Not Provided"}<br/>
                     <strong>Place:</strong> ${p.City || "Unknown City"}, ${p.CountyName || "Unknown"} County<br/>
                     <strong>Road:</strong> ${p.Road || "Not Provided"}<br/>
                     <strong>Direction:</strong> ${direction}<br/>
                     <strong>Reason:</strong> ${p.Reason || "Not Provided"}<br/>
                     <strong>Until:</strong> ${p.EndDateET || "Not Provided"}`;
      (layer as L.Path).bindTooltip(label, { sticky: true });
      layer.on("mouseover", () => (layer as L.Path).setStyle({ color: "yellow" }));
      layer.on("mouseout", () => (layer as L.Path).setStyle({ color: "blue" }));
    };

    // -------- Work zones --------
    const tims = "https://services.arcgis.com/NuWFvHYDMVmmxMeM/ArcGIS/rest/services/NCDOT_TIMSIncidentsByIncidentType/FeatureServer/1";
    const wz = (where: string) => LEsri.featureLayer({ url: tims, where, pane: "hccs", style: workzoneStyle, onEachFeature: workzoneEachFeature } as any) as EsriFeatureLayer;
    this.truckClosureLayer = wz("IncidentType = 'Truck Closure'");
    this.constructionLayer = wz("IncidentType = 'Construction'");
    this.nightConstructionLayer = wz("IncidentType = 'Night Time Construction'");
    this.maintenanceLayer = wz("IncidentType = 'Maintenance'");
    this.nightMaintenanceLayer = wz("IncidentType = 'Night Time Maintenance'");
    this.emergencyLayer = wz("IncidentType = 'Emergency Road Work'");
    this.obstructionLayer = wz("IncidentType = 'Road Obstruction'");
    this.weatherLayer = wz("IncidentType = 'Weather Event'");
    this.specialLayer = wz("IncidentType = 'Special Event'");
    this.otherLayer = wz("IncidentType = 'Other'");

    // -------- Boundaries --------
    const boundaryStyle: L.PathOptions = { fill: true, fillOpacity: 0.0, fillColor: "#000", weight: 1, color: "gray", dashArray: "3" };
    this.troopLayer = LEsri.featureLayer({
      url: "https://ags.coverlab.org/server/rest/services/Basedata/Boundaries/FeatureServer/3",
      pane: "counties", style: boundaryStyle,
      onEachFeature: (_f: GeoJSON.Feature, layer: L.Layer) => {
        const p: any = (_f as any).properties || {};
        (layer as L.Path).bindTooltip(`Troop ${p.Troop ?? "Troop"}`, { sticky: true });
        layer.on("mouseover", () => (layer as L.Path).setStyle({ weight: 3, color: "yellow" }));
        layer.on("mouseout", () => (layer as L.Path).setStyle({ weight: 1, color: "gray" }));
      }
    } as any) as EsriFeatureLayer;

    this.mpoLayer = LEsri.featureLayer({
      url: "https://ags.coverlab.org/server/rest/services/Basedata/Boundaries/FeatureServer/0",
      pane: "counties", style: boundaryStyle,
      onEachFeature: (_f: GeoJSON.Feature, layer: L.Layer) => {
        const p: any = (_f as any).properties || {};
        (layer as L.Path).bindTooltip(String(p.Name ?? "MPO"), { sticky: true });
        layer.on("mouseover", () => (layer as L.Path).setStyle({ weight: 3, color: "yellow" }));
        layer.on("mouseout", () => (layer as L.Path).setStyle({ weight: 1, color: "gray" }));
      }
    } as any) as EsriFeatureLayer;

    this.lelRegionsLayer = LEsri.featureLayer({
      url: "https://ags.coverlab.org/server/rest/services/Basedata/Boundaries/FeatureServer/2",
      pane: "lel", style: boundaryStyle,
      onEachFeature: (_f: GeoJSON.Feature, layer: L.Layer) => {
        const p: any = (_f as any).properties || {};
        (layer as L.Path).bindTooltip(`LEL Region ${p.LEL_REGION ?? "LEL Region"}`, { sticky: true });
        layer.on("mouseover", () => (layer as L.Path).setStyle({ weight: 3, color: "yellow" }));
        layer.on("mouseout", () => (layer as L.Path).setStyle({ weight: 1, color: "gray" }));
      }
    } as any) as EsriFeatureLayer;

    // -------- HCCS layer (default LEL mode statewide) --------
    this.currentMode = "LEL";
    this.lastWhere = "LELRank IN (1,2,3,4,5)";

    this.hccsStyle = (feature: GeoJSON.Feature) => {
      const p = feature.properties as any;
      const rank = this.currentMode === "LEL" ? p?.LELRank : p?.CntyRankVeh;
      const { color, weight } = this.getRankStyle(Number(rank));
      return { color, weight, opacity: 0.9, interactive: true };
    };

    this.hccsLayer = LEsri.featureLayer({
      url: "https://ags.coverlab.org/server/rest/services/HighCrashCorridors/HCCs/FeatureServer/0",
      where: this.lastWhere, pane: "hccs",
      idField: "OBJECTID",
      fields: ["OBJECTID", "RouteName", "LEL_Region", "Troop", "County", "LELRank", "CntyRankVeh", "CrashDateStart", "CrashDateEnd"],
      simplifyFactor: 0.5, precision: 5,
      style: this.hccsStyle,
      onEachFeature: (feature: GeoJSON.Feature, layer: L.Layer) => {
        const path = layer as L.Path;
        path.bindTooltip("", { sticky: true, direction: "top" });
        const html = () => {
          const p = feature.properties as any;
          const rank = this.currentMode === "LEL" ? p?.LELRank : p?.CntyRankVeh;
          const label = this.currentMode === "LEL" ? "LEL Rank:" : "County Rank:";
          return `<strong>Route:</strong> ${p.RouteName || "Unknown"}<br/>
                  <strong>LEL:</strong> ${p.LEL_Region || "Unknown"}<br/>
                  <strong>Troop:</strong> ${p.Troop || "Unknown"}<br/>
                  <strong>County:</strong> ${p.County || "Unknown"}<br/>
                  <strong>${label}</strong> ${rank ?? "Unknown"}<br/>
                  <strong>Range of Data:</strong> ${p.CrashDateStart} - ${p.CrashDateEnd}`;
        };
        path.on("tooltipopen", () => path.setTooltipContent(html()));
        path.on("mouseover", () => { path.setStyle({ weight: 20 }); path.setTooltipContent(html()); path.openTooltip(); });
        path.on("mouseout", () => {
          const p = feature.properties as any;
          const r = this.currentMode === "LEL" ? p?.LELRank : p?.CntyRankVeh;
          const { weight } = this.getRankStyle(Number(r));
          path.setStyle({ weight });
        });
      }
    } as any) as EsriFeatureLayer;

    this.hccsLayer.addTo(this.map);

    // call these in your constructor after layers are created:
    this.startAutoRefresh(6 * 60 * 60 * 1000);  // every 6 hours (tune as you like)
    document.addEventListener("visibilitychange", this.onVisibilityRefresh);

    // Retry on 400 (e.g., LEL_Region type mismatch)
    this.hccsLayer.on("requesterror", (e: any) => {
      console.warn("[HCCS requesterror]", e, "WHERE:", e?.params?.where);
      if (e?.params?.where === this.lastWhere && this.lastLelAltWhere && this.hccsLayer) {
        const alt = this.lastLelAltWhere;
        this.lastLelAltWhere = null;
        this.lastWhere = alt;
        (this.hccsLayer as EsriFeatureLayer).setWhere(alt);
        (this.hccsLayer as any).refresh?.();
        console.warn("[HCCS] Retrying with alternate WHERE:", alt);
      }
    });

    // Initial fit
    this.hccsLayer.once("load", () => {
      this.hccsLayer!.query().where(this.lastWhere).bounds((err, b) => {
        if (!err && b && b.isValid()) this.map.fitBounds(b.pad(0.05));
      });
    });

    // -------- Layer control --------
    const baseMaps = { "Dark Basemap": this.darkMap, "Light Basemap": this.lightMap };
    const groupedOverlays: any = {
      "Boundaries": {
        "Counties": this.countiesLayer,
        "LEL Regions": this.lelRegionsLayer!,
        "Planning Orgs": this.mpoLayer!,
        "Troops": this.troopLayer!
      },
      "Work Zones": {
        "Truck Closure": this.truckClosureLayer!,
        "Construction": this.constructionLayer!,
        "Night Construction": this.nightConstructionLayer!,
        "Maintenance": this.maintenanceLayer!,
        "Night Time Maintenance": this.nightMaintenanceLayer!,
        "Emergency Road Work": this.emergencyLayer!,
        "Road Obstruction": this.obstructionLayer!,
        "Weather Event": this.weatherLayer!,
        "Special Event": this.specialLayer!,
        "Other": this.otherLayer!
      },
      "Crash Data": { "High Crash Corridors": this.hccsLayer! }
    };

    // @ts-ignore
    this.layerControl = (L as any).control.groupedLayers(
      baseMaps, groupedOverlays,
      { collapsed: false, exclusiveGroups: ["Boundaries"], groupCheckboxes: false }
    ).addTo(this.map) as L.Control.Layers;

    this.map.on("overlayadd overlayremove baselayerchange", () => {
      requestAnimationFrame(() => this.refreshLayerControlCollapsibles());
    });

    this.refreshLayerControlCollapsibles();
    this.setBoundaryModeStable(this.currentMode);

    // -------- Legend --------
    const Legend = L.Control.extend({
      options: { position: "bottomleft" },
      onAdd: () => {
        const div = L.DomUtil.create("div", "leaflet-control legend");
        div.innerHTML = `<h4>Crash Rank</h4>`;
        [1, 2, 3, 4, 5].forEach(r => {
          const { color, weight } = this.getRankStyle(r);
          const h = Math.max(weight + 6, 14); // room for thick lines
          div.innerHTML += `
        <span style="display:inline-block;width:36px;height:${h}px;position:relative;margin-right:8px;vertical-align:middle;">
          <span style="
            position:absolute;left:0;right:0;top:50%;
            height:0;border-top:${weight}px solid ${color};
            transform:translateY(-50%);
          "></span>
        </span>
        <span style="vertical-align:middle;">${r}</span><br/>`;
        });
        return div;
      }
    });
    this.legendControl = new Legend();
    this.legendControl.addTo(this.map);
  }

  public update(options: VisualUpdateOptions): void {
    const dv = options.dataViews?.[0];
    console.log("[update] fired", {
      hasDV: !!dv,
      categories: dv?.categorical?.categories?.length ?? 0,
      values: dv?.categorical?.values?.length ?? 0,
      hasTable: !!dv?.table,
      rows: dv?.table?.rows?.length ?? 0
    });

    this.resizeMap(options);

    if (this.pendingFitTimer !== null) {
      clearTimeout(this.pendingFitTimer);
      this.pendingFitTimer = null;
    }
    if (!dv) return;

    // helpers
    const norm = (s: string) => String(s ?? "")
      .replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim()
      .replace(/\s+COUNTY$/i, "").toUpperCase();
    const onlyDigits = (s: string) => String(s ?? "").replace(/\D+/g, "");
    const esc = (s: string) => s.replace(/'/g, "''");

    // role values
    const lelRegionsRaw = this.getRoleValues(dv, "lelRegion");
    const countiesRaw = this.getRoleValues(dv, "county");
    const troopsRaw = this.getRoleValues(dv, "troop");

    const selectedLelRegions = new Set(lelRegionsRaw.map(v => String(v).trim().toUpperCase()));

    const selectedCountyBases = new Set<string>();
    for (const raw of countiesRaw) {
      const v = String(raw).trim();
      const digits = onlyDigits(v);
      if ((/^\d{3}$/.test(digits) || /^\d{5}$/.test(digits)) && (fipsToCounty as any)[Number(digits)]) {
        selectedCountyBases.add(norm((fipsToCounty as any)[Number(digits)]));
      } else {
        selectedCountyBases.add(norm(v));
      }
    }

    const selectedTroops = new Set(
      troopsRaw.map(t => String(t).trim().toUpperCase().replace(/^TROOP\s+/, ""))
    );

    // keys used to detect *changes* across updates
    const lelKey = Array.from(selectedLelRegions).sort().join("|");
    const countyKey = Array.from(selectedCountyBases).sort().join("|");
    const troopKey = Array.from(selectedTroops).sort().join("|");

    // measures (keep your reads)
    const countyModeFlag = this.getSingleMeasure(dv, "countyModeFlag") > 0;
    const lelModeFlag = this.getSingleMeasure(dv, "lelModeFlag") > 0;

    // detect what changed since last time
    const countyChanged = (countyKey !== this.lastCountyKey);
    const lelChanged = (lelKey !== this.lastLelKey);
    const troopChanged = (troopKey !== this.lastTroopKey);

    // a gentle fallback if user picked exactly one county and didn't touch LEL
    const singleCountyFallback =
      (selectedCountyBases.size === 1) && (selectedLelRegions.size === 0 || lelRegionsRaw.length === 0);

    // mode selection priority:
    // 1) explicit flags  2) change-based inference  3) single-county fallback  4) hold current mode
    let desiredMode: "LEL" | "COUNTY";
    if (countyModeFlag) desiredMode = "COUNTY";
    else if (lelModeFlag) desiredMode = "LEL";
    else if (countyChanged && !lelChanged) desiredMode = "COUNTY";
    else if (lelChanged && !countyChanged) desiredMode = "LEL";
    else if (troopChanged && !lelChanged) desiredMode = "COUNTY";
    else if (singleCountyFallback) desiredMode = "COUNTY";
    else if (selectedTroops.size > 0 && !lelChanged && !countyChanged) desiredMode = "COUNTY";
    else desiredMode = this.currentMode;

    // WHERE
    // replace your WHERE block in update(...) with this:

    let computedWhere: string;

    if (desiredMode === "COUNTY") {
      this.lastLelAltWhere = null;

      const allowCountyTokens = true;
      const allowTroopTokens = selectedTroops.size > 0;

      const bases = allowCountyTokens ? Array.from(selectedCountyBases) : [];
      const countyList = bases.length
        ? bases.map(b => [`'${esc(b)}'`, `'${esc(b + " COUNTY")}'`]).flat().join(",")
        : "";

      const troopTokens: string[] = [];
      if (allowTroopTokens) {
        for (const t of selectedTroops) {
          const tt = t.replace(/'/g, "''");
          troopTokens.push(`'${tt}'`, `'TROOP ${tt}'`);
        }
      }
      const troopList = troopTokens.length ? Array.from(new Set(troopTokens)).join(",") : "";

      if (countyList && troopList) {
        computedWhere = `(UPPER(County) IN (${countyList}) AND UPPER(Troop) IN (${troopList}) AND CntyRankVeh IN (1,2,3,4,5))`;
      } else if (countyList) {
        computedWhere = `(UPPER(County) IN (${countyList}) AND CntyRankVeh IN (1,2,3,4,5))`;
      } else if (troopList) {
        computedWhere = `(UPPER(Troop) IN (${troopList}) AND CntyRankVeh IN (1,2,3,4,5))`;
      } else {
        computedWhere = `CntyRankVeh IN (1,2,3,4,5)`;
      }
    } else {
      computedWhere = this.buildLelWhere(selectedLelRegions);
      // Accept both base and "NAME COUNTY" variants while in LEL mode
      if (selectedCountyBases.size) {
        const countyTokens = Array.from(selectedCountyBases)
          .map(n => [`'${esc(n)}'`, `'${esc(n + " COUNTY")}'`])
          .flat();
        const countyVals = Array.from(new Set(countyTokens)).join(",");
        computedWhere = `(${computedWhere}) AND UPPER(County) IN (${countyVals})`;
      }
    }

    // one call only (new signature)
    this.scheduleApply(desiredMode, computedWhere, lelKey, countyKey, troopKey);

    // tidy, accurate debug info
    console.log("[HCCS debug update]", {
      lelRegionsRaw, countiesRaw, troopsRaw,
      countySelected: Array.from(selectedCountyBases),
      troopSelected: Array.from(selectedTroops),
      countyModeFlag, lelModeFlag,
      desiredMode,
      usingCountyTokens: desiredMode === "COUNTY" && selectedCountyBases.size > 0,
      usingTroopTokens: desiredMode === "COUNTY" && selectedTroops.size > 0,
      where: computedWhere
    });

  }

  public destroy(): void {
    if (this.autoRefreshTimer) clearInterval(this.autoRefreshTimer);
    document.removeEventListener("visibilitychange", this.onVisibilityRefresh);
    this.map.remove();
  }

  // ---------------- internals ----------------

  private createMapContainer() {
    const existing = document.getElementById("mapid");
    if (existing) existing.remove();
    const div = document.createElement("div");
    div.id = "mapid";
    div.style.width = "100%";
    div.style.height = "100%";
    this.target.appendChild(div);
    this.mapContainer = div;
  }

  private initMap() {
    this.map = L.map("mapid", { center: [35.54, -79.24], zoom: 7, maxZoom: 20, minZoom: 3 });

    this.map.createPane("counties"); this.map.getPane("counties")!.style.zIndex = "400";
    this.map.createPane("lel"); this.map.getPane("lel")!.style.zIndex = "500";
    this.map.createPane("hccs"); this.map.getPane("hccs")!.style.zIndex = "600";

    this.darkMap = L.tileLayer("https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png", {
      attribution: '&copy; <a href="https://carto.com/">CARTO</a>', subdomains: "abcd", maxZoom: 20
    }).addTo(this.map);

    this.lightMap = L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19, attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    });
  }

  private resizeMap(opts: VisualUpdateOptions) {
    requestAnimationFrame(() => {
      this.mapContainer.style.width = `${opts.viewport.width}px`;
      this.mapContainer.style.height = `${opts.viewport.height}px`;
      this.map.invalidateSize();
    });
  }

  private getRankStyle(rank: number): { color: string; weight: number } {
    switch (rank) {
      case 5: return { color: "#ffff00", weight: 2 };
      case 4: return { color: "#ff9900", weight: 4 };
      case 3: return { color: "#ff0000", weight: 6 };
      case 2: return { color: "#990000", weight: 8 };
      case 1: return { color: "#660000", weight: 10 };
      default: return { color: "rgba(255, 255, 255, 0)", weight: 0 };
    }
  }

  private makeGroupedLayerCollapsible(control: L.Control, groupName: string) {
    const container = (control as any).getContainer?.() as HTMLElement ?? (control as any)._container as HTMLElement;
    if (!container) return;
    const nameSpan = Array.from(container.querySelectorAll(".leaflet-control-layers-group-name"))
      .find(el => (el.textContent ?? "").trim().toLowerCase() === groupName.trim().toLowerCase()) as HTMLElement | undefined;
    if (!nameSpan) return;
    const headerLabel = (nameSpan.closest(".leaflet-control-layers-group-label") as HTMLElement) || nameSpan;
    const groupRoot = (nameSpan.closest(".leaflet-control-layers-group") as HTMLElement) || headerLabel.parentElement!;
    if (!groupRoot) return;

    const hdr = headerLabel as HTMLElement & { _gliWired?: boolean };
    if (hdr._gliWired) return; hdr._gliWired = true;

    const items = Array.from(groupRoot.children).filter(el => el !== headerLabel) as HTMLElement[];
    if (!headerLabel.querySelector(".gli-caret")) {
      const caret = document.createElement("span");
      caret.className = "gli-caret"; caret.setAttribute("role", "button");
      caret.setAttribute("aria-label", "Toggle group"); caret.tabIndex = 0;
      headerLabel.insertBefore(caret, nameSpan);
    }
    const setCollapsed = (collapsed: boolean) => {
      items.forEach(el => { el.style.display = collapsed ? "none" : ""; });
      headerLabel.classList.toggle("is-collapsed", collapsed);
    };
    const toggle = () => setCollapsed(!(items[0]?.style.display === "none"));

    headerLabel.style.cursor = "pointer";
    headerLabel.addEventListener("click", (e: MouseEvent) => {
      e.preventDefault(); e.stopPropagation(); toggle();
    });
    headerLabel.addEventListener("keydown", (e: KeyboardEvent) => {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggle(); }
    });

    L.DomEvent.disableClickPropagation(headerLabel);
    L.DomEvent.disableScrollPropagation(headerLabel);

    setCollapsed(true);
  }

  private refreshLayerControlCollapsibles() {
    if (!this.layerControl) return;
    requestAnimationFrame(() => {
      this.makeGroupedLayerCollapsible(this.layerControl!, "Boundaries");
      this.makeGroupedLayerCollapsible(this.layerControl!, "Work Zones");
      this.makeGroupedLayerCollapsible(this.layerControl!, "Crash Data");
    });
  }

  private setBoundaryModeStable(mode: "LEL" | "COUNTY") {
    const want = mode === "COUNTY" ? this.countiesLayer : this.lelRegionsLayer;

    // Everything that lives in the "Boundaries" group
    const boundaryLayers: (L.Layer | undefined)[] = [
      this.countiesLayer,
      this.lelRegionsLayer,
      this.mpoLayer,
      this.troopLayer
    ];

    // If nothing to do, bail
    const wantOn = !!want && this.map.hasLayer(want as any);
    const othersOn = boundaryLayers.some(l => l && l !== want && this.map.hasLayer(l as any));
    if (wantOn && !othersOn) return;

    const refresh = () => requestAnimationFrame(() => this.refreshLayerControlCollapsibles());
    this.map.once("layeradd", refresh);
    this.map.once("layerremove", refresh);

    // Turn OFF everything except the desired boundary layer
    for (const l of boundaryLayers) {
      if (l && l !== want && this.map.hasLayer(l as any)) {
        this.map.removeLayer(l as any);
      }
    }

    // Ensure the desired boundary layer is ON
    if (want && !this.map.hasLayer(want as any)) {
      this.map.addLayer(want as any);
    }
  }

  // Role readers
  private getRoleValues(dv: powerbi.DataView | undefined, role: string): string[] {
    if (!dv) return [];
    const seen = new Set<string>();
    const push = (v: any) => { if (v !== null && v !== undefined) seen.add(String(v)); };
    const c = dv.categorical;
    if (c?.categories?.length) for (const col of c.categories) if (col.source.roles?.[role]) for (const v of col.values) push(v);
    if (c?.values?.length) for (const col of c.values) if (col.source.roles?.[role]) for (const v of col.values as any[]) push(v);
    return Array.from(seen);
  }

  private getSingleMeasure(dv: powerbi.DataView | undefined, role: string): number {
    if (!dv) return 0;
    const c = dv.categorical;
    const vcol = c?.values?.find(v => !!v.source.roles?.[role]);
    if (vcol) {
      const n = Number(vcol.values?.[0]);
      return Number.isFinite(n) ? n : 0;
    }
    return 0;
  }

  // Fit to current WHERE (server-side)
  private debouncedFitToHccs(token: number, delayMs = 350) {
    if (this.pendingFitTimer) clearTimeout(this.pendingFitTimer);
    this.pendingFitTimer = window.setTimeout(() => {
      if (token !== this.updateToken) return;
      const lyr = this.hccsLayer as EsriFeatureLayer;
      if (!lyr) return;
      lyr.query().where(this.lastWhere).bounds((err, b) => {
        if (token !== this.updateToken) return;
        if (!err && b && b.isValid()) this.map.fitBounds(b.pad(0.05));
      });
    }, delayMs);
  }

  // was: private scheduledApply = false;
  //      private queuedState?: { desiredMode: "LEL" | "COUNTY"; where: string; lelKey: string };

  private scheduledApply = false;
  private queuedState?: {
    desiredMode: "LEL" | "COUNTY";
    where: string;
    lelKey: string;
    countyKey: string;
    troopKey: string;
  };

  private scheduleApply(desiredMode: "LEL" | "COUNTY", where: string, lelKey: string, countyKey: string, troopKey: string) {
    this.queuedState = { desiredMode, where, lelKey, countyKey, troopKey };
    if (this.scheduledApply) return;
    this.scheduledApply = true;
    requestAnimationFrame(() => {
      this.scheduledApply = false;
      const s = this.queuedState!;
      this.queuedState = undefined;
      this.applyState(s);
    });
  }

  private applyState(s: { desiredMode: "LEL" | "COUNTY"; where: string; lelKey: string; countyKey: string; troopKey: string }) {
    const layer = this.hccsLayer as any;
    if (!layer) return;

    const token = ++this.updateToken;
    const modeChanged = (this.currentMode !== s.desiredMode);
    const whereChanged = (this.lastWhere !== s.where);

    this.currentMode = s.desiredMode;

    if (whereChanged) {
      this.lastWhere = s.where;
      layer.setWhere(s.where);
      layer.refresh?.();
    }
    if (modeChanged) layer.setStyle(this.hccsStyle);

    this.setBoundaryModeStable(this.currentMode);
    this.debouncedFitToHccs(token);

    // remember latest selections
    this.lastLelKey = s.lelKey;
    this.lastCountyKey = s.countyKey;
    this.lastTroopKey = s.troopKey;
  }

  // PATCH A — replace entire method
  private buildLelWhere(selectedLelRegions: Set<string>): string {
    // LEL_Region is TEXT: e.g., "1","2","3"
    if (selectedLelRegions.size === 0) {
      this.lastLelAltWhere = null;
      return "LELRank IN (1,2,3,4,5)";
    }

    const esc = (s: string) => s.replace(/'/g, "''");

    // pull out the region number if the slicer label contains words (e.g., "Region 3")
    const vals = Array.from(selectedLelRegions)
      .map(s => (String(s).match(/\d+/)?.[0] ?? String(s)).trim())
      .filter(s => s.length > 0)
      .map(s => `'${esc(s)}'`);

    const inList = Array.from(new Set(vals)).join(",");
    this.lastLelAltWhere = null; // no alternate WHERE needed anymore
    return `LEL_Region IN (${inList}) AND LELRank IN (1,2,3,4,5)`;
  }

  // helpers
  private onVisibilityRefresh = () => {
    if (document.visibilityState === "visible") this.refreshAllLayers();
  };

  private startAutoRefresh(ms: number) {
    if (this.autoRefreshTimer) clearInterval(this.autoRefreshTimer);
    if (ms > 0) this.autoRefreshTimer = window.setInterval(() => this.refreshAllLayers(), ms);
  }

  private refreshAllLayers() {
    const layers = [
      this.hccsLayer,
      this.truckClosureLayer, this.constructionLayer, this.nightConstructionLayer,
      this.maintenanceLayer, this.nightMaintenanceLayer, this.emergencyLayer,
      this.obstructionLayer, this.weatherLayer, this.specialLayer, this.otherLayer
    ];
    for (const l of layers) (l as any)?.refresh?.();
  }

}
