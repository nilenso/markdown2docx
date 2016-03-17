(ns markdown2docx.docx-writer
  (:require [markdown2docx.docx :as docx]))

;; `visit` is forward-declared because it consumes all the leaf fns which themselves
;; recursively call `visit`
(declare visit)

;; Maitinging atoms as the entire namespace deals explicitly with a mutable object and changes need to cascade laterally, not just down the tree.
(defonce numid (atom 1))
(defonce ilvl (atom -1))

(defn reset-list
  [ndp]
  (when (= 0 @ilvl)
    (swap! numid (docx/restart-numbering ndp)))
  (swap! ilvl dec))

(defn document
  [content maindoc parent]
  (let [traversed (conj parent :document)]
    (doseq [child content]
     ((visit child) maindoc traversed))))

(defn heading
  [content maindoc parent ilvl]
  (let [lvl (-> content
                first
                :lvl)
        p (docx/add-paragraph maindoc)
        traversed (conj parent :heading)]
    (docx/add-style-heading p lvl)
    (doseq [child  (rest content)]
      ((visit child) p traversed ilvl))))

(defn paragraph
  [content maindoc parent]
  (let [p (docx/add-paragraph maindoc)
        traversed (conj parent :paragraph)]
    (doseq [child content]
      ((visit child) p traversed))))

(defn table
  [content maindoc parent]
  (let [table (docx/add-table maindoc)
        traversed (conj parent :table)]
    (doseq [child content]
      ((visit child) table traversed))))

(defn table-head
  [content table parent]
  (let [traversed (conj parent :table-head)]
    (doseq [child content]
      ((visit child) table traversed))))

(defn table-body
  [content table parent]
  (let [traversed (conj parent :table-body)]
    (doseq [child content]
      ((visit child) table traversed))))

(defn table-row
  [content table parent]
  (let [row (docx/add-table-row table)
        traversed (conj parent :table-row)]
    (doseq [child content]
      ((visit child) row traversed))))

(defn table-cell
  [content row parent]
  (let [cell (docx/add-table-cell row)
        traversed (conj parent :table-cell)]
    (doseq [child content]
      ((visit child) cell traversed))))

(defn ordered-list
  [content maindoc parent]
  (let [ndp (docx/add-ordered-list maindoc)
        traversed (conj parent :ordered-list)]
    (swap! ilvl inc)
    (doseq [child content]
      ((visit child) maindoc traversed))
    (reset-list ndp)))

(defn bullet-list
  [content maindoc parent]
  (let [traversed (conj parent :bullet-list)]
    (doseq [child content]
      ((visit child) maindoc traversed))))

(defn list-item
  [content maindoc parent]
  (let [traversed (conj parent :list-item)]
    (doseq [child content]
      ((visit child) maindoc traversed))))

(defn hard-line-break
  [content p parent]
  (let [traversed (conj parent :hard-line-break)]
    (doseq [child content]
      ((visit child) p traversed))))

(defn thematic-break
  [content p parent]
  (let [traversed (conj parent :thematic-break)]
   (doseq [child content]
     ((visit child) p traversed))))

(defn soft-line-break
  [content p parent]
  (let [traversed (conj parent :soft-line-break)]
    (doseq [child content]
      ((visit child) p traversed))))

(defn bold
  [content p parent]
  (docx/add-emphasis-text p :bold (-> content
                                      first
                                      :text)))

(defn italic
  [content p parent]
  (docx/add-emphasis-text p :italic (-> content
                                        first
                                        :text)))

(defn text
  [content p parent]
  (println parent)
  (if (= :list-item (second parent))
    (docx/add-text-list p @numid @ilvl content)
    (docx/add-text p content)))

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
  [maindoc document-map]
  (let [root-node '(:package)]
    ((visit document-map) maindoc root-node)))

(defn write
  [docx-file document-map]
  (let [package (docx/create-package)
        maindoc (docx/maindoc package)
        footer-part (docx/create-footer-part package)]
    (docx/create-footer-reference package footer-part)
    (build-docx maindoc document-map)
    (docx/save package docx-file)))
