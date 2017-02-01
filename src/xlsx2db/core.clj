(ns xlsx2db.core
  (:require [dk.ative.docjure.spreadsheet :as s]
            [clojure.string :refer [join]]
            [clojure.java.jdbc :as j]
            [clojure.data :as d]
  #_(:import [org.apache.poi.ss.usermodel Cell])
  (:gen-class)))

;; Database
(def hsql-db {:subprotocol "hsqldb"
              :subname "file:/tmp/tramwaydb/db"
              :user "SA"
              :password ""})

(def quote-fn (j/quoted \"))

(defn- ->sql-columns [columns quoted]
  (map #(str  (quoted %) " varchar(255)") columns))

(defn- table-ddl [table-name columns quoted]
  (str "create text table " (quoted  table-name) " (" (join ", " (->sql-columns columns quoted)) ");"))

(defn- text-table-dll [table-name quoted]
  (str "SET TABLE " (quoted table-name) " SOURCE \""  table-name ".csv\";"))

(defn create-table
  [hsql-db table-name columns quoted]
  (j/db-do-commands hsql-db (table-ddl table-name columns quoted))
  (j/db-do-commands hsql-db (text-table-dll table-name quoted)))

;; spreadshhet
(defn- column-headers [sheet]
  "Returns the values of the first row in the sheet"
  (->> sheet s/row-seq first s/cell-seq (map s/read-cell)))

(defn- sheet-name [sheet] (.getSheetName sheet))

(defn- column-name-seq []
  (map #(keyword (str (char %))) (range 65 91)))

(defn- headers-map
  "ToDo Only works for 26 columns"
  [sheet]
  (zipmap (column-name-seq) (column-headers sheet)))

;; Conversion xlsx -> db
(defn- sheet->table
  [hsql-db sheet quoted]
  (create-table hsql-db (sheet-name sheet) (column-headers sheet) quoted))

(defn- sheet-data [sheet]
  (rest (s/select-columns (headers-map sheet) sheet)))

(defn sheetdata->table
  [hsql-db sheet quoted]
  (j/insert-multi! hsql-db (sheet-name sheet) (sheet-data sheet) {:entities quoted}))

(defn wb->db [wb hsql-db quoted]
  (j/with-db-connection [db hsql-db]
    (doseq [s (s/sheet-seq wb)]
      (sheet->table db s quoted)
      (sheetdata->table db s quoted))))

(defn row-data [map keys] (mapv #(get map %) keys))

(defn maps->sheet-data [maps]
  (let [head (apply vector (keys (first maps)))]
    (conj (map #(row-data % head) maps) head)))

;; ToDo allow multiple sheets instead of sheet
(defn create-wb
  [path sheet]
  (let [out (s/create-workbook (:name sheet) (maps->sheet-data (:rows sheet)))]
    (s/save-workbook! path out)))

;(def wb (s/load-workbook-from-file "/tmp/SampleData.xlsx"))
;(def sheet (s/select-sheet "SalesOrders" wb))
;(def data (sheet-data sheet))
;(wb->db wb hsql-db quote-fn)
;(create-wb "/tmp/arst1.xlsx" {:name "mysheet" :rows data})

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  (println "Hello, World!"))
