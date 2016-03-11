(ns markdown2docx.core
  (:require [clojure.java.io :refer [file]]
            [clojure.string :as string]
            [markdown2docx.docx :as docx]
            [markdown2docx.md :as md]))

(declare visit)
(defonce numid (atom 1))
(defonce ilvl (atom -1))

(defn adjust-list-ind-level
  [ndp]
  (when (= 0 @ilvl)
    (swap! numid (docx/restart-numbering ndp)))
  (swap! ilvl dec))

(defn document
  [content maindoc]
  (doseq [child content]
    ((visit child) maindoc)))

(defn heading
  [content maindoc]
  (let [lvl (-> content
                first
                :lvl)
        p (docx/add-paragraph maindoc)]
       (docx/add-style-heading p lvl)
       (doseq [child  (rest content)]
         ((visit child) p))))

(defn paragraph
  [content maindoc & parent]
  (let [p (docx/add-paragraph maindoc)]
    (doseq [child content]
      ((visit child) p parent))))

(defn table
  [content maindoc]
  (let [table (docx/add-table maindoc)]
    (doseq [child content]
      ((visit child) table))))

(defn table-head
  [content table]
  (doseq [child content]
    ((visit child) table)))

(defn table-body
  [content table]
  (doseq [child content]
    ((visit child) table)))

(defn table-row
  [content table]
  (let [row (docx/add-table-row table)]
    (doseq [child content]
      ((visit child) row))))

(defn table-cell
  [content row]
  (let [cell (docx/add-table-cell row)]
    (doseq [child content]
      ((visit child) cell))))

(defn ordered-list
  [content maindoc & parent]
  (let [ndp (docx/add-ordered-list maindoc)]
    (swap! ilvl inc)
    (doseq [child content]
      ((visit child) maindoc))
    (adjust-list-ind-level ndp)))

(defn bullet-list
  [content maindoc & parent]
  (doseq [child content]
    ((visit child) maindoc)))

(defn list-item
  [content maindoc & parent]
  (doseq [child content]
    ((visit child) maindoc :list-item)))

(defn hard-line-break
  [content p & parent]
  (doseq [child content]
    ((visit child) p)))

(defn thematic-break
  [content p & parent]
  (doseq [child content]
    ((visit child) p)))

(defn soft-line-break
  [content p & parent]
  (doseq [child content]
    ((visit child) p)))

(defn bold
  [content p & parent]
  (docx/add-emphasis-text p :bold (-> content
                                      first
                                      :text)))


(defn italic
  [content p & parent]
  (docx/add-emphasis-text p :italic (-> content
                                        first
                                        :text)))

(defn text
  [content p & parent]
  (if (nil? parent)
    (docx/add-text p content)
    (docx/add-text-list p @numid @ilvl content)))

(defn visit
  [element]
  (let [key (first (keys element))
        value (first (vals element))]
    (case key
      :document (partial document value)
      :heading (partial heading value)
      :paragraph (partial paragraph value)
      :table-block (partial table value)
      :table-head (partial table-head value)
      :table-body (partial table-body value)
      :table-row (partial table-row value)
      :table-cell (partial table-cell value)
      :ordered-list (partial ordered-list value)
      :bullet-list (partial bullet-list value)
      :list-item (partial list-item value)
      :hard-line-break (partial hard-line-break value)
      :thematic-break (partial thematic-break value)
      :soft-line-break (partial soft-line-break value)
      :bold (partial bold value)
      :italic (partial italic value)
      :text (partial text value))))

(defn build-docx
  [maindoc md-map]
  ((visit md-map) maindoc))

(defn md2docx
  [docx-file md-map]
  (let [package (docx/create-package)
        maindoc (docx/maindoc package)
        footer-part (docx/create-footer-part package)]
    (docx/create-footer-reference package footer-part)
    (build-docx maindoc md-map)
    (docx/save package docx-file)))

(defn -main [& [md-file docx-file]]
  (let [md-string (slurp md-file)
        md-map (md/parse md-string)]
    (md2docx docx-file md-map)))
