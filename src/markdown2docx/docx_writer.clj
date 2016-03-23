(ns markdown2docx.docx-writer
  (:require [markdown2docx.docx :as docx]))

;; `visit` is forward-declared because it consumes all the leaf fns which themselves
;; recursively call `visit`
(declare visit)

;; Maitinging atoms as the entire namespace deals explicitly with a mutable object and changes need to cascade laterally, not just down the tree.
(defonce numid (atom 1))
(defonce ilvl (atom -1))

(defn in?
  "true if coll contains elm"
  [coll elm]
  (some #(= elm %) coll))

(defn remove-css-element
  [text]
  (clojure.string/replace text #"\{([^\}]+)\}$" ""))

(defn check-css-tags
  [text]
  (-> (->> text
           (re-find #"\{([^\}]+)\}$"))
      (or [])
      last
      keyword))

(defn get-rules
  [text css-map]
  (->> text
       check-css-tags
       (get css-map)))

(defn get-rule-keys
  [rules]
  (keys rules))

(defn css-bold
  [rule-ks]
  (in? rule-ks :font-weight))

(defn css-italic
  [rule-ks]
  (in? rule-ks :font-style))

(defn css-bold-and-italics
  [rule-ks]
  (and (css-bold rule-ks) (css-italic rule-ks)))

(defn css-bold-or-italics
  [rule-ks]
  (or (css-bold rule-ks) (css-italic rule-ks)))

(defn css-alignment
  [rule-ks]
  (in? rule-ks :text-align))

(defn apply-css
  [p text css-map]
  (let [rules (get-rules text css-map)
        rule-keys (get-rule-keys rules)]
    (when (css-alignment rule-keys)
      (docx/add-text-align p (keyword (:text-align rules))))
    (when (css-bold-and-italics rule-keys)
      (docx/add-bold-italic-text p (remove-css-element text)))
    (when (and (css-bold rule-keys) (not (css-bold-and-italics rule-keys)))
      (docx/add-emphasis-text p :bold (remove-css-element text)))
    (when (and (css-italic rule-keys) (not (css-bold-and-italics rule-keys)))
      (docx/add-emphasis-text p :italic (remove-css-element text)))
    (remove-css-element text)))

(defn reset-list
  [ndp]
  (when (= 0 @ilvl)
    (swap! numid (docx/restart-numbering ndp)))
  (swap! ilvl dec))

(defn document
  [content maindoc parent css-map]
  (let [traversed (conj parent :document)]
    (doseq [child content]
     ((visit child) maindoc traversed css-map))))

(defn heading
  [content maindoc parent css-map]
  (let [lvl (-> content
                first
                :level)
        p (docx/add-paragraph maindoc)
        traversed (conj parent :heading)]
    (docx/add-style-heading p lvl)
    (doseq [child  (rest content)]
      ((visit child) p traversed css-map))))

(defn paragraph
  [content maindoc parent css-map]
  (let [p (docx/add-paragraph maindoc)
        traversed (conj parent :paragraph)]
    (doseq [child content]
      ((visit child) p traversed css-map))))

(defn table
  [content maindoc parent css-map]
  (let [table (docx/add-table maindoc)
        traversed (conj parent :table)]
    (doseq [child content]
      ((visit child) table traversed css-map))))

(defn table-head
  [content table parent css-map]
  (let [traversed (conj parent :table-head)]
    (doseq [child content]
      ((visit child) table traversed css-map))))

(defn table-body
  [content table parent css-map]
  (let [traversed (conj parent :table-body)]
    (doseq [child content]
      ((visit child) table traversed css-map))))

(defn table-row
  [content table parent css-map]
  (let [row (docx/add-table-row table)
        traversed (conj parent :table-row)]
    (doseq [child content]
      ((visit child) row traversed css-map))))

(defn table-cell
  [content row parent css-map]
  (let [cell (docx/add-table-cell row)
        traversed (conj parent :table-cell)]
    (doseq [child content]
      ((visit child) cell traversed css-map))))

(defn ordered-list
  [content maindoc parent css-map]
  (let [ndp (docx/add-ordered-list maindoc)
        traversed (conj parent :ordered-list)]
    (swap! ilvl inc)
    (doseq [child content]
      ((visit child) maindoc traversed css-map))
    (reset-list ndp)))

(defn bullet-list
  [content maindoc parent css-map]
  (let [traversed (conj parent :bullet-list)]
    (doseq [child content]
      ((visit child) maindoc traversed css-map))))

(defn list-item
  [content maindoc parent css-map]
  (let [traversed (conj parent :list-item)]
    (doseq [child content]
      ((visit child) maindoc traversed css-map))))

(defn hard-line-break
  [content p parent css-map]
  (let [traversed (conj parent :hard-line-break)]
    (doseq [child content]
      ((visit child) p traversed css-map))))

(defn thematic-break
  [content p parent css-map]
  (let [traversed (conj parent :thematic-break)]
   (doseq [child content]
     ((visit child) p traversed css-map))))

(defn soft-line-break
  [content p parent css-map]
  (let [traversed (conj parent :soft-line-break)]
    (doseq [child content]
      ((visit child) p traversed css-map))))

(defn bold
  [content p parent css-map]
  (docx/add-emphasis-text p :bold (-> content
                                      first
                                      :text)))

(defn italic
  [content p parent css-map]
  (docx/add-emphasis-text p :italic (-> content
                                        first
                                        :text)))

(defn text
  [content p parent css-map]
  (if (= :list-item (second parent))
    (docx/add-text-list p @numid @ilvl content)
    (if (-> content
            (get-rules css-map)
            (get-rule-keys)
            (css-bold-or-italics))
      (apply-css p content css-map)
      (docx/add-text p (apply-css p content css-map)))))

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
  [maindoc document-map css-map]
  (let [root-node '(:package)]
    ((visit document-map) maindoc root-node css-map)))

(defn write
  [docx-file document-map css-map]
  (let [package (docx/create-package)
        maindoc (docx/maindoc package)
        footer-part (docx/create-footer-part package)]
    (docx/create-footer-reference package footer-part)
    (build-docx maindoc document-map css-map)
    (docx/save package docx-file)))
