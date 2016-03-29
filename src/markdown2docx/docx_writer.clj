(ns markdown2docx.docx-writer
  (:require [markdown2docx.docx :as docx]))

;; `visit` is forward-declared because it consumes all the leaf fns which themselves
;; recursively call `visit`
(declare visit)

;; Maitinging atoms as the entire namespace deals explicitly with a mutable object and changes need to cascade laterally, not just down the tree.
(defonce numid (atom 1))
(defonce ilvl (atom -1))

(defn in?
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
  [doc text css-map]
  (let [rules (get-rules text css-map)
        rule-keys (get-rule-keys rules)
        align (when (css-alignment rule-keys)
                (docx/add-text-align doc (keyword (:text-align rules))))
        bold-italic (when (css-bold-and-italics rule-keys)
                      (docx/add-bold-italic-text doc))
        bold (when (and (css-bold rule-keys) (not (css-bold-and-italics rule-keys)))
               (docx/set-emphasis-text doc :bold ))
        italic (when (and (css-italic rule-keys) (not (css-bold-and-italics rule-keys)))
                 (docx/set-emphasis-text doc :italic))]
    (or bold italic bold-italic align doc)))

(defn reset-list
  [ndp]
  (when (= 0 @ilvl)
    (swap! numid inc))
  (swap! ilvl dec))

(defn document
  [content doc parent css-map child-num]
  (let [traversed (conj parent :document)]
    (doseq [child content]
     ((visit child) doc traversed css-map (.indexOf content child)))))

(defn heading
  [content doc parent css-map child-num]
  (let [lvl (-> content
                first
                :level)
        new-doc (docx/add-paragraph doc)
        traversed (conj parent :heading)]
    (docx/add-style-heading new-doc lvl)
    (doseq [child  (rest content)]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn paragraph
  [content doc parent css-map child-num]
  (let [new-doc (if (and (= :list-item (first parent)) (= 0 child-num))
                  doc
                  (docx/add-paragraph doc))
        traversed (conj parent :paragraph)]
    (when (and (= :list-item (first parent)) (not= 0 child-num))
      (docx/indent-paragraph new-doc @ilvl))
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn table
  [content doc parent css-map child-num]
  (let [new-doc (docx/add-table doc)
        traversed (conj parent :table)]
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn table-head
  [content doc parent css-map child-num]
  (let [traversed (conj parent :table-head)]
    (doseq [child content]
      ((visit child) doc traversed css-map (.indexOf content child)))))

(defn table-body
  [content doc parent css-map child-num]
  (let [traversed (conj parent :table-body)]
    (doseq [child content]
      ((visit child) doc traversed css-map (.indexOf content child)))))

(defn table-row
  [content doc parent css-map child-num]
  (let [new-doc (docx/add-table-row doc)
        traversed (conj parent :table-row)]
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn table-cell
  [content doc parent css-map child-num]
  (let [new-doc (docx/add-table-cell doc)
        traversed (conj parent :table-cell)]
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn ordered-list
  [content doc parent css-map child-num]
  (let [new-doc (docx/add-ordered-list doc)
        traversed (conj parent :ordered-list)]
    (swap! ilvl inc)
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))
    (reset-list (:ndp new-doc))))

(defn bullet-list
  [content doc parent css-map child-num]
  (let [traversed (conj parent :bullet-list)]
    (doseq [child content]
      ((visit child) doc traversed css-map (.indexOf content child)))))

(defn list-item
  [content doc parent css-map child-num]
  (let [new-doc (docx/add-paragraph doc)
        traversed (conj parent :list-item)]
    (docx/add-text-list new-doc @numid @ilvl)
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn indented-code-block
  [content doc parent css-map child-num]
  (let [new-doc (docx/add-paragraph doc)
        traversed (conj parent :indented-code-block)]
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn page-break
  [content doc parent css-map child-num]
  (docx/add-page-break doc))

(defn hard-line-break
  [content doc parent css-map child-num]
  (let [traversed (conj parent :hard-line-break)]
    (doseq [child content]
      ((visit child) doc traversed css-map (.indexOf content child)))))

(defn thematic-break
  [content doc parent css-map child-num]
  (let [traversed (conj parent :thematic-break)]
    (doseq [child content]
     ((visit child) doc traversed css-map (.indexOf content child)))))

(defn soft-line-break
  [content doc parent css-map child-num]
  (let [traversed (conj parent :soft-line-break)]
    (doseq [child content]
      ((visit child) doc traversed css-map (.indexOf content child)))))

(defn bold
  [content doc parent css-map child-num]
  (let [new-doc (docx/set-emphasis-text doc :bold)
        traversed (conj parent :bold)]
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn italic
  [content doc parent css-map child-num]
  (let [new-doc (docx/set-emphasis-text doc :italic)
        traversed (conj parent :italic)]
    (doseq [child content]
      ((visit child) new-doc traversed css-map (.indexOf content child)))))

(defn text
  [content doc parent css-map child-num]
  (docx/add-text (apply-css doc content css-map) (remove-css-element content)))

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
      :page-break (partial page-break value)
      :hard-line-break (partial hard-line-break value)
      :thematic-break (partial thematic-break value)
      :soft-line-break (partial soft-line-break value)
      :indented-code-block (partial indented-code-block value)
      :bold (partial bold value)
      :italic (partial italic value)
      :text (partial text value))))

(defn build-docx
  [maindoc document-map css-map]
  (let [root-node '(:package)]
    ((visit document-map) {:maindoc maindoc} root-node css-map 1)))

(defn write
  [docx-file document-map css-map]
  (let [package (docx/create-package)
        maindoc (docx/maindoc package)
        footer-part (docx/create-footer-part package)]
    (docx/create-footer-reference package footer-part)
    (build-docx maindoc document-map css-map)
    (docx/save package docx-file)))
