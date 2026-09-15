import { useMemo, useRef, useState } from 'react'
import './App.css'

type ExcelHeader = {
  columnIndex: number
  columnLetter: string
  name: string
  sampleValues: string[]
}

type LegendEntry = {
  color: string
  label: string
}

type ExcelInspection = {
  worksheets: string[]
  worksheet: string
  headerRow: number
  totalDataRows: number
  headers: ExcelHeader[]
  legend: LegendEntry[]
}

type CoordinateMode = 'MissingOnly' | 'All'
type RouteType = 'TODAS' | 'VENTA' | 'ENTREGA'

const routeTypeOptions: { value: RouteType; title: string; description: string }[] = [
  { value: 'TODAS', title: 'Ventas y entregas', description: 'Todas las sucursales, un marcador por cada una.' },
  { value: 'VENTA', title: 'Solo ventas', description: 'Direcciones de venta; si no hay, usa las de entrega.' },
  { value: 'ENTREGA', title: 'Solo entregas', description: 'Solo las direcciones de entrega del cliente.' },
]
type CoordinateLayout = 'None' | 'Combined' | 'Separate'

type ExportSummary = {
  rows: number
  placemarks: number
  requested: number
  found: number
  updated: number
  unresolved: number
  notFound: number
  repeated: number
  extraSites: number
}

const clientCodePattern = /(c[oó]d(igo)?\s*(de\s*)?cliente|cliente\s*(id|c[oó]digo)|customer\s*(id|code))/i
const clientNamePattern = /(nombre[\s_-]*(del[\s_-]*)?cliente|cliente[\s_-]*nombre|customer[\s_-]*name|raz[oó]n[\s_-]*social|^nombre$|^name$)/i
const latitudePattern = /^(lat|latitud|latitude|coord(enada)?\s*y)$/i
const longitudePattern = /^(lon|lng|longitud|longitude|coord(enada)?\s*x)$/i
const combinedCoordinatesPattern = /^(coordenadas?|coordinates?|lat(itud)?\s*[,/+-]\s*(lon|lng|longitud)|ubicaci[oó]n)$/i

function App() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [file, setFile] = useState<File | null>(null)
  const [headerRow, setHeaderRow] = useState(1)
  const [inspection, setInspection] = useState<ExcelInspection | null>(null)
  const [selectedColumns, setSelectedColumns] = useState<Set<number>>(new Set())
  const [clientCodeColumn, setClientCodeColumn] = useState<number | null>(null)
  const [clientNameColumn, setClientNameColumn] = useState<number | null>(null)
  const [latitudeColumn, setLatitudeColumn] = useState<number | null>(null)
  const [longitudeColumn, setLongitudeColumn] = useState<number | null>(null)
  const [coordinatesColumn, setCoordinatesColumn] = useState<number | null>(null)
  const [coordinateLayout, setCoordinateLayout] = useState<CoordinateLayout>('None')
  const [coordinateMode, setCoordinateMode] = useState<CoordinateMode>('MissingOnly')
  const [routeType, setRouteType] = useState<RouteType>('TODAS')
  const [useLegend, setUseLegend] = useState(false)
  const [useRowColors, setUseRowColors] = useState(true)
  // La leyenda se relaciona por color, por eso requiere los colores de las filas.
  const canUseLegend = useRowColors && (inspection?.legend.length ?? 0) > 0
  const [isDragging, setIsDragging] = useState(false)
  const [isInspecting, setIsInspecting] = useState(false)
  const [isExporting, setIsExporting] = useState(false)
  const [error, setError] = useState('')
  const [summary, setSummary] = useState<ExportSummary | null>(null)

  const hasValidCoordinateSelection =
    coordinateLayout === 'None' ||
    (coordinateLayout === 'Combined' && coordinatesColumn !== null) ||
    (coordinateLayout === 'Separate' && latitudeColumn !== null && longitudeColumn !== null)
  const canExport = Boolean(
    file &&
      inspection &&
      selectedColumns.size > 0 &&
      clientCodeColumn &&
      clientNameColumn &&
      hasValidCoordinateSelection &&
      !isExporting,
  )

  const progressStep = inspection ? 3 : file ? 2 : 1
  const selectedNames = useMemo(
    () =>
      inspection?.headers
        .filter((header) => selectedColumns.has(header.columnIndex))
        .map((header) => header.name) ?? [],
    [inspection, selectedColumns],
  )

  const resetInspection = () => {
    setInspection(null)
    setSelectedColumns(new Set())
    setClientCodeColumn(null)
    setClientNameColumn(null)
    setLatitudeColumn(null)
    setLongitudeColumn(null)
    setCoordinatesColumn(null)
    setCoordinateLayout('None')
    setUseLegend(false)
    setSummary(null)
  }

  const selectFile = (nextFile: File | null) => {
    setError('')
    if (!nextFile) return

    if (!/\.(xlsx|xlsm)$/i.test(nextFile.name)) {
      setError('Selecciona un archivo Excel con extensión .xlsx o .xlsm.')
      return
    }

    setFile(nextFile)
    resetInspection()
  }

  const inspectFile = async (worksheet?: string) => {
    if (!file) {
      setError('Selecciona un archivo Excel antes de continuar.')
      return
    }

    setIsInspecting(true)
    setError('')
    setSummary(null)

    try {
      const form = new FormData()
      form.append('File', file)
      form.append('HeaderRow', String(headerRow))
      if (worksheet) form.append('Worksheet', worksheet)

      const response = await fetch('/api/excel/inspect', { method: 'POST', body: form })
      if (!response.ok) throw new Error(await readApiError(response))

      const result = (await response.json()) as ExcelInspection
      setInspection(result)
      setSelectedColumns(new Set(result.headers.map((header) => header.columnIndex)))
      setUseLegend(result.legend.length > 0)
      setClientCodeColumn(findHeader(result.headers, clientCodePattern))
      setClientNameColumn(findHeader(result.headers, clientNamePattern))
      const combinedColumn = findHeader(result.headers, combinedCoordinatesPattern)
      const detectedLatitude = findHeader(result.headers, latitudePattern)
      const detectedLongitude = findHeader(result.headers, longitudePattern)
      setCoordinatesColumn(combinedColumn)
      setLatitudeColumn(detectedLatitude)
      setLongitudeColumn(detectedLongitude)
      setCoordinateLayout(
        combinedColumn !== null
          ? 'Combined'
          : detectedLatitude !== null || detectedLongitude !== null
            ? 'Separate'
            : 'None',
      )
    } catch (requestError) {
      setError(messageFrom(requestError))
    } finally {
      setIsInspecting(false)
    }
  }

  const exportCsv = async () => {
    if (!file || !inspection || !clientCodeColumn || !clientNameColumn) return

    setIsExporting(true)
    setError('')
    setSummary(null)

    try {
      const form = new FormData()
      form.append('File', file)
      form.append('HeaderRow', String(headerRow))
      form.append('Worksheet', inspection.worksheet)
      Array.from(selectedColumns)
        .sort((a, b) => a - b)
        .forEach((column) => form.append('SelectedColumns', String(column)))
      form.append('ClientCodeColumn', String(clientCodeColumn))
      form.append('ClientNameColumn', String(clientNameColumn))
      if (coordinateLayout === 'Separate' && latitudeColumn) form.append('LatitudeColumn', String(latitudeColumn))
      if (coordinateLayout === 'Separate' && longitudeColumn) form.append('LongitudeColumn', String(longitudeColumn))
      if (coordinateLayout === 'Combined' && coordinatesColumn) form.append('CoordinatesColumn', String(coordinatesColumn))
      form.append('CoordinateMode', coordinateMode)
      form.append('RouteType', routeType)
      form.append('UseRowColors', String(useRowColors))
      form.append('UseLegend', String(canUseLegend && useLegend))

      const response = await fetch('/api/excel/export', { method: 'POST', body: form })
      if (!response.ok) throw new Error(await readApiError(response))

      const blob = await response.blob()
      downloadBlob(blob, getDownloadName(response, file.name))
      setSummary({
        rows: headerNumber(response, 'X-Rows-Exported'),
        placemarks: headerNumber(response, 'X-Kml-Placemarks'),
        requested: headerNumber(response, 'X-Coordinates-Requested'),
        found: headerNumber(response, 'X-Coordinates-Found'),
        updated: headerNumber(response, 'X-Rows-Updated'),
        unresolved: headerNumber(response, 'X-Rows-Unresolved'),
        notFound: headerNumber(response, 'X-Rows-Not-Found'),
        repeated: headerNumber(response, 'X-Codes-Repeated'),
        extraSites: headerNumber(response, 'X-Sites-Extra'),
      })
    } catch (requestError) {
      setError(messageFrom(requestError))
    } finally {
      setIsExporting(false)
    }
  }

  const toggleColumn = (columnIndex: number) => {
    setSelectedColumns((current) => {
      const next = new Set(current)
      if (next.has(columnIndex)) next.delete(columnIndex)
      else next.add(columnIndex)
      return next
    })
  }

  const toggleAll = () => {
    if (!inspection) return
    const allSelected = selectedColumns.size === inspection.headers.length
    setSelectedColumns(
      allSelected ? new Set() : new Set(inspection.headers.map((header) => header.columnIndex)),
    )
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#top" aria-label="MyMaps CSV Builder">
          <span className="brand-mark" aria-hidden="true">
            <MapPinIcon />
          </span>
          <span>
            <strong>MyMaps</strong>
            <small>CSV Builder</small>
          </span>
        </a>
        <div className="privacy-pill">
          <ShieldIcon />
          Tus archivos se procesan de forma segura
        </div>
      </header>

      <section className="hero" id="top">
        <div className="eyebrow">EXCEL → GOOGLE MY MAPS</div>
        <h1>Convierte tus clientes en un mapa.</h1>
        <p>
          Elige los datos que necesitas, completa las coordenadas y descarga un paquete listo para
          importar.
        </p>

        <ol className="steps" aria-label={`Paso actual ${progressStep} de 3`}>
          <li className={progressStep >= 1 ? 'active' : ''}>
            <span>1</span>
            <small className="step-full">Cargar archivo</small>
            <small className="step-short">Cargar</small>
          </li>
          <li className={progressStep >= 2 ? 'active' : ''}>
            <span>2</span>
            <small className="step-full">Configurar datos</small>
            <small className="step-short">Configurar</small>
          </li>
          <li className={progressStep >= 3 ? 'active' : ''}>
            <span>3</span>
            <small className="step-full">Generar archivos</small>
            <small className="step-short">Generar</small>
          </li>
        </ol>
      </section>

      <section className="workspace">
        <article className="card upload-card">
          <div className="card-heading">
            <span className="step-icon"><UploadIcon /></span>
            <div>
              <span className="section-label">PASO 1</span>
              <h2>Carga tu archivo Excel</h2>
              <p>Archivos .xlsx o .xlsm de hasta 50 MB</p>
            </div>
          </div>

          <button
            type="button"
            className={`dropzone ${isDragging ? 'dragging' : ''} ${file ? 'has-file' : ''}`}
            onClick={() => inputRef.current?.click()}
            onDragEnter={(event) => { event.preventDefault(); setIsDragging(true) }}
            onDragOver={(event) => event.preventDefault()}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(event) => {
              event.preventDefault()
              setIsDragging(false)
              selectFile(event.dataTransfer.files[0] ?? null)
            }}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xlsm"
              onChange={(event) => selectFile(event.target.files?.[0] ?? null)}
            />
            <span className="upload-illustration"><FileIcon /></span>
            {file ? (
              <>
                <strong>{file.name}</strong>
                <span>{formatBytes(file.size)} · Haz clic para cambiarlo</span>
              </>
            ) : (
              <>
                <strong>Arrastra tu archivo aquí</strong>
                <span>o haz clic para seleccionarlo</span>
              </>
            )}
          </button>

          <div className="inspect-controls">
            <label>
              ¿En qué fila están los encabezados?
              <span className="input-with-icon">
                <RowsIcon />
                <input
                  type="number"
                  min="1"
                  max="1048576"
                  value={headerRow}
                  onChange={(event) => {
                    setHeaderRow(Math.max(1, Number(event.target.value) || 1))
                    resetInspection()
                  }}
                />
              </span>
            </label>
            <button
              className="button secondary"
              type="button"
              disabled={!file || isInspecting}
              onClick={() => inspectFile()}
            >
              {isInspecting ? <Spinner /> : <SearchIcon />}
              {isInspecting ? 'Leyendo Excel…' : 'Mostrar encabezados'}
            </button>
          </div>
        </article>

        {error && (
          <div className="alert error" role="alert">
            <AlertIcon />
            <div><strong>No pudimos completar la operación</strong><span>{error}</span></div>
          </div>
        )}

        {inspection && (
          <>
            <article className="card configure-card">
              <div className="card-heading split-heading">
                <div className="heading-group">
                  <span className="step-icon violet"><SlidersIcon /></span>
                  <div>
                    <span className="section-label">PASO 2</span>
                    <h2>Configura los datos</h2>
                    <p>Selecciona las columnas que aparecerán en el CSV.</p>
                  </div>
                </div>
                <div className="workbook-meta">
                  {inspection.worksheets.length > 1 ? (
                    <label>
                      Hoja
                      <select
                        value={inspection.worksheet}
                        disabled={isInspecting}
                        onChange={(event) => inspectFile(event.target.value)}
                      >
                        {inspection.worksheets.map((worksheet) => (
                          <option key={worksheet}>{worksheet}</option>
                        ))}
                      </select>
                    </label>
                  ) : (
                    <span><SheetIcon /> {inspection.worksheet}</span>
                  )}
                  <span>{inspection.totalDataRows.toLocaleString('es-DO')} filas de datos</span>
                </div>
              </div>

              <div className="table-toolbar">
                <div>
                  <strong>{selectedColumns.size}</strong> de {inspection.headers.length} columnas seleccionadas
                </div>
                <button type="button" className="text-button" onClick={toggleAll}>
                  {selectedColumns.size === inspection.headers.length ? 'Deseleccionar todas' : 'Seleccionar todas'}
                </button>
              </div>

              <div className="headers-table" role="table" aria-label="Encabezados del Excel">
                <div className="header-row table-head" role="row">
                  <span role="columnheader">Incluir</span>
                  <span role="columnheader">Columna del Excel</span>
                  <span role="columnheader">Muestra de datos</span>
                  <span role="columnheader">Uso especial</span>
                </div>
                {inspection.headers.map((header) => {
                  const badges = [
                    clientCodeColumn === header.columnIndex ? 'Código de cliente' : '',
                    clientNameColumn === header.columnIndex ? 'Nombre del cliente' : '',
                    latitudeColumn === header.columnIndex ? 'Latitud' : '',
                    longitudeColumn === header.columnIndex ? 'Longitud' : '',
                    coordinatesColumn === header.columnIndex ? 'Latitud + longitud' : '',
                  ].filter(Boolean)

                  return (
                    <div className="header-row" role="row" key={header.columnIndex}>
                      <span role="cell">
                        <label className="check-control" aria-label={`Incluir ${header.name}`}>
                          <input type="checkbox" checked={selectedColumns.has(header.columnIndex)} onChange={() => toggleColumn(header.columnIndex)} />
                          <span><CheckIcon /></span>
                        </label>
                      </span>
                      <span className="column-name" role="cell">
                        <b>{header.columnLetter}</b>
                        <span>{header.name}</span>
                      </span>
                      <span className="samples" role="cell">
                        {header.sampleValues.length ? header.sampleValues.join(' · ') : <em>Sin datos de muestra</em>}
                      </span>
                      <span className="badges" role="cell">
                        {badges.length ? badges.map((badge) => <small key={badge}>{badge}</small>) : <span className="muted">—</span>}
                      </span>
                    </div>
                  )
                })}
              </div>

              <div className="special-columns">
                <div className="subheading">
                  <h3>Identifica las columnas clave</h3>
                  <p>El código permite buscar cada cliente y el nombre aparecerá en la descripción del marcador.</p>
                </div>
                <div className="select-grid base-settings">
                  <ColumnSelect label="Código del cliente" value={clientCodeColumn} headers={inspection.headers} required onChange={setClientCodeColumn} />
                  <ColumnSelect label="Nombre del cliente" value={clientNameColumn} headers={inspection.headers} required onChange={setClientNameColumn} />
                  <label className="column-select">
                    <span>Formato de coordenadas</span>
                    <select
                      value={coordinateLayout}
                      onChange={(event) => {
                        const layout = event.target.value as CoordinateLayout
                        setCoordinateLayout(layout)
                        if (layout !== 'Combined') setCoordinatesColumn(null)
                        if (layout !== 'Separate') {
                          setLatitudeColumn(null)
                          setLongitudeColumn(null)
                        }
                      }}
                    >
                      <option value="None">No están en el Excel</option>
                      <option value="Combined">Latitud y longitud en una columna</option>
                      <option value="Separate">Latitud y longitud en columnas separadas</option>
                    </select>
                  </label>
                </div>
                {coordinateLayout === 'Combined' && (
                  <div className="select-grid coordinate-pickers">
                    <ColumnSelect label="Columna con latitud y longitud" value={coordinatesColumn} headers={inspection.headers} required onChange={setCoordinatesColumn} />
                    <p className="coordinate-example">
                      <b>Formato esperado</b>
                      <code>19.793532,-70.703465</code>
                      Se convertirá en dos columnas para el CSV y el KML.
                    </p>
                  </div>
                )}
                {coordinateLayout === 'Separate' && (
                  <div className="select-grid coordinate-pickers">
                    <ColumnSelect label="Latitud existente" value={latitudeColumn} headers={inspection.headers} required onChange={setLatitudeColumn} />
                    <ColumnSelect label="Longitud existente" value={longitudeColumn} headers={inspection.headers} required onChange={setLongitudeColumn} />
                  </div>
                )}
                {!hasValidCoordinateSelection && <p className="field-warning">Selecciona las columnas necesarias para el formato elegido.</p>}
              </div>

              <div className="special-columns legend-settings">
                <div className="subheading">
                  <h3>Colores y leyenda</h3>
                  <p>Elige si el mapa usa el color de relleno de cada fila y la leyenda del Excel.</p>
                </div>
                <div className="legend-options">
                  <label className="legend-toggle">
                    <span className="check-control">
                      <input
                        type="checkbox"
                        checked={useRowColors}
                        onChange={(event) => setUseRowColors(event.target.checked)}
                      />
                      <span><CheckIcon /></span>
                    </span>
                    <span>
                      <strong>Cargar el color de cada fila</strong>
                      <small>Agrega la columna Color y pinta cada marcador. Si lo desactivas, todos usan el azul de My Maps.</small>
                    </span>
                  </label>
                  <label className={`legend-toggle ${canUseLegend ? '' : 'disabled'}`}>
                    <span className="check-control">
                      <input
                        type="checkbox"
                        checked={canUseLegend && useLegend}
                        disabled={!canUseLegend}
                        onChange={(event) => setUseLegend(event.target.checked)}
                      />
                      <span><CheckIcon /></span>
                    </span>
                    <span>
                      <strong>Usar la leyenda como columna Categoría</strong>
                      <small>
                        {inspection.legend.length === 0
                          ? 'No encontramos celdas con color y texto arriba de la fila de encabezados.'
                          : useRowColors
                            ? `Encontramos ${inspection.legend.length} colores con texto arriba de la fila de encabezados.`
                            : 'Activa el color de cada fila para usar la leyenda.'}
                      </small>
                    </span>
                  </label>
                </div>
                {inspection.legend.length > 0 && (
                  <ul className={`legend-list ${canUseLegend && useLegend ? '' : 'inactive'}`} aria-label="Leyenda encontrada">
                    {inspection.legend.map((entry) => (
                      <li key={entry.color}>
                        <span className="legend-swatch" style={{ background: entry.color }} aria-hidden="true" />
                        {entry.label}
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </article>

            <article className="card coordinate-card">
              <div className="card-heading">
                <span className="step-icon coral"><CrosshairIcon /></span>
                <div>
                  <span className="section-label">COORDENADAS</span>
                  <h2>¿Qué coordenadas deseas consultar?</h2>
                  <p>El API procesa los códigos en grupos de hasta 2,000.</p>
                </div>
              </div>
              <div className="mode-grid">
                <ModeOption checked={coordinateMode === 'MissingOnly'} title="Solo completar las faltantes" description="Conserva las coordenadas válidas del Excel y consulta solo donde falten." icon={<SparkIcon />} onClick={() => setCoordinateMode('MissingOnly')} />
                <ModeOption checked={coordinateMode === 'All'} title="Consultar y validar todas" description="Consulta cada código y reemplaza las coordenadas cuando el API tenga datos." icon={<RefreshIcon />} onClick={() => setCoordinateMode('All')} />
              </div>
              <div className="route-type-panel">
                <div className="subheading">
                  <h3>¿Qué direcciones del GPS consultar?</h3>
                  <p>Un cliente puede tener una sucursal de venta y varias de entrega; cada una sale como marcador.</p>
                </div>
                <div className="route-type-grid" role="radiogroup" aria-label="Direcciones del GPS">
                  {routeTypeOptions.map((option) => (
                    <button
                      key={option.value}
                      type="button"
                      role="radio"
                      aria-checked={routeType === option.value}
                      className={`route-type-option ${routeType === option.value ? 'selected' : ''}`}
                      onClick={() => setRouteType(option.value)}
                    >
                      <span className="radio-dot" aria-hidden="true" />
                      <span><strong>{option.title}</strong><small>{option.description}</small></span>
                    </button>
                  ))}
                </div>
              </div>
              <div className="color-note">
                <span className="color-swatch" />
                Importa el KML o KMZ en My Maps y deja la capa en Estilos individuales para conservar el color de relleno de cada fila. Si agrupas por Color, My Maps asigna sus propios colores a cada grupo. Si tu Excel tiene una leyenda con colores arriba de los encabezados, se agrega la columna Categoría: agrupa por ella para usar el panel de la capa como leyenda.
              </div>
            </article>

            <article className="export-panel">
              <div>
                <span className="section-label">PASO 3</span>
                <h2>Tu paquete está listo para generarse</h2>
                <p>Incluye CSV, KML y KMZ con colores exactos · {selectedNames.length} columnas · {inspection.totalDataRows.toLocaleString('es-DO')} filas</p>
              </div>
              <button className="button primary" type="button" disabled={!canExport} onClick={exportCsv}>
                {isExporting ? <Spinner /> : <DownloadIcon />}
                {isExporting ? 'Consultando y generando…' : 'Descargar CSV + KML + KMZ'}
              </button>
            </article>
          </>
        )}

        {summary && (
          <div className="alert success" role="status">
            <CheckCircleIcon />
            <div>
              <strong>Paquete CSV + KML + KMZ descargado correctamente</strong>
              <span>{summary.rows} filas ubicables · {summary.placemarks} marcadores · importa el KMZ para conservar los colores · {summary.requested} códigos consultados · {summary.found} encontrados · {summary.updated} filas actualizadas{summary.unresolved > 0 ? ` · ${summary.unresolved} sin ubicación en el mapa` : ''}{summary.notFound > 0 ? ` · ${summary.notFound} en el CSV de no encontrados` : ''}{summary.repeated > 0 ? ` · ${summary.repeated} con coordenadas repetidas en GPS (ver CSV)` : ''}{summary.extraSites > 0 ? ` · ${summary.extraSites} marcadores de sucursales adicionales` : ''}</span>
            </div>
          </div>
        )}
      </section>

      <footer><MapPinIcon /> CSV + KML + KMZ preparados para Google My Maps</footer>
    </main>
  )
}

function ColumnSelect({ label, value, headers, required = false, onChange }: {
  label: string
  value: number | null
  headers: ExcelHeader[]
  required?: boolean
  onChange: (value: number | null) => void
}) {
  return (
    <label className="column-select">
      <span>{label}{required && <b>REQUERIDO</b>}</span>
      <select value={value ?? ''} onChange={(event) => onChange(event.target.value ? Number(event.target.value) : null)}>
        <option value="">{required ? 'Seleccionar columna…' : 'No está en el Excel'}</option>
        {headers.map((header) => <option key={header.columnIndex} value={header.columnIndex}>{header.columnLetter} — {header.name}</option>)}
      </select>
    </label>
  )
}

function ModeOption({ checked, title, description, icon, onClick }: {
  checked: boolean
  title: string
  description: string
  icon: React.ReactNode
  onClick: () => void
}) {
  return (
    <button type="button" className={`mode-option ${checked ? 'selected' : ''}`} onClick={onClick}>
      <span className="mode-icon">{icon}</span>
      <span><strong>{title}</strong><small>{description}</small></span>
      <span className="radio-dot" aria-hidden="true" />
    </button>
  )
}

function findHeader(headers: ExcelHeader[], pattern: RegExp) {
  return headers.find((header) => pattern.test(header.name.trim()))?.columnIndex ?? null
}

async function readApiError(response: Response) {
  try {
    const problem = (await response.json()) as { detail?: string; title?: string; errors?: Record<string, string[]> }
    const validation = problem.errors ? Object.values(problem.errors).flat().join(' ') : ''
    return problem.detail || validation || problem.title || `Error ${response.status}`
  } catch {
    return `El servidor respondió con el error ${response.status}.`
  }
}

function messageFrom(error: unknown) {
  return error instanceof Error ? error.message : 'Ocurrió un error inesperado.'
}

function headerNumber(response: Response, name: string) {
  return Number(response.headers.get(name) ?? 0)
}

function getDownloadName(response: Response, originalName: string) {
  const disposition = response.headers.get('Content-Disposition') ?? ''
  const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/i)?.[1]
  const regular = disposition.match(/filename="?([^";]+)"?/i)?.[1]
  if (encoded) return decodeURIComponent(encoded)
  if (regular) return regular
  return `${originalName.replace(/\.[^.]+$/, '')}-google-mymaps.zip`
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  URL.revokeObjectURL(url)
}

function formatBytes(bytes: number) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 ** 2).toFixed(1)} MB`
}

const Icon = ({ children }: { children: React.ReactNode }) => <svg viewBox="0 0 24 24" aria-hidden="true">{children}</svg>
const MapPinIcon = () => <Icon><path d="M20 10c0 5-8 12-8 12S4 15 4 10a8 8 0 1 1 16 0Z"/><circle cx="12" cy="10" r="2.5"/></Icon>
const ShieldIcon = () => <Icon><path d="M12 3 5 6v5c0 4.8 2.9 8 7 10 4.1-2 7-5.2 7-10V6l-7-3Z"/><path d="m9 12 2 2 4-4"/></Icon>
const UploadIcon = () => <Icon><path d="M12 16V4m0 0L7 9m5-5 5 5"/><path d="M4 15v5h16v-5"/></Icon>
const FileIcon = () => <Icon><path d="M6 2h8l4 4v16H6z"/><path d="M14 2v5h5M9 13h6M9 17h6"/></Icon>
const RowsIcon = () => <Icon><path d="M4 6h16M4 12h16M4 18h16"/><path d="M8 4v16"/></Icon>
const SearchIcon = () => <Icon><circle cx="10.5" cy="10.5" r="6.5"/><path d="m16 16 5 5"/></Icon>
const SlidersIcon = () => <Icon><path d="M4 7h10M18 7h2M4 17h2M10 17h10"/><circle cx="16" cy="7" r="2"/><circle cx="8" cy="17" r="2"/></Icon>
const SheetIcon = () => <Icon><path d="M5 3h14v18H5zM5 9h14M10 9v12"/></Icon>
const CheckIcon = () => <Icon><path d="m5 12 4 4L19 6"/></Icon>
const CrosshairIcon = () => <Icon><circle cx="12" cy="12" r="7"/><circle cx="12" cy="12" r="2"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3"/></Icon>
const SparkIcon = () => <Icon><path d="m12 3 1.5 5.5L19 10l-5.5 1.5L12 17l-1.5-5.5L5 10l5.5-1.5L12 3Z"/><path d="m19 16 .6 2.4L22 19l-2.4.6L19 22l-.6-2.4L16 19l2.4-.6L19 16Z"/></Icon>
const RefreshIcon = () => <Icon><path d="M20 6v5h-5"/><path d="M18.5 8A8 8 0 1 0 20 15"/></Icon>
const DownloadIcon = () => <Icon><path d="M12 3v13m0 0 5-5m-5 5-5-5"/><path d="M4 20h16"/></Icon>
const AlertIcon = () => <Icon><path d="M12 3 2 21h20L12 3Z"/><path d="M12 9v5M12 18h.01"/></Icon>
const CheckCircleIcon = () => <Icon><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></Icon>
const Spinner = () => <span className="spinner" aria-hidden="true" />

export default App
