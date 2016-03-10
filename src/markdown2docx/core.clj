(ns markdown2docx.core
  (:require [clojure.java.io :refer [file]]
            [clojure.string :as string]
            [markdown2docx.docx :as docx])
  (:import [com.steadystate.css.parser CSSOMParser]
           [org.w3c.dom.css CSSStyleSheet CSSRuleList CSSRule CSSStyleRule
            CSSStyleDeclaration]
           [org.w3c.css.sac InputSource]
           [java.io StringReader]
           [org.commonmark.parser Parser]
           [org.commonmark.html HtmlRenderer]
           [org.commonmark Extension]
           [org.commonmark.ext.gfm.tables TablesExtension TableBlock TableCell
            TableBody TableHead TableRow]
           [org.commonmark.node Visitor AbstractVisitor Block BlockQuote
            BulletList Code CustomBlock CustomNode Document
            Emphasis FencedCodeBlock HardLineBreak Heading
            HtmlBlock HtmlInline Image IndentedCodeBlock Link
            ListBlock ListItem Node OrderedList Paragraph
            SoftLineBreak StrongEmphasis Text ThematicBreak]))

;; Temp files
(def md-file "/home/sp4de/Nilenso/cooperative-agreement/template.md")
(def docx-file "/home/sp4de/Nilenso/cooperative-agreement/template.docx")
(def css-file "/home/sp4de/Nilenso/cooperative-agreement/style.css")
(def stack (atom '()))
(def package (docx/create-package))
(def maindoc (docx/maindoc package))
(def numid (atom 1))
(def ilvl (atom -1))

(defn push
  [stack val]
  (swap! stack conj val))

(defn pop!!
  [stack]
  (swap! stack rest))

(defn head
  [stack]
  (first @stack))

(defn property-as-map [property]
 (let [split-property (-> property .toString (clojure.string/split #": "))]
   {(-> split-property
        (get 0)
        keyword)
    (-> split-property
        (get 1))}))

(defn rule-as-map [rule]
  (let [k  (-> rule
               .getSelectors
               .getSelectors
               (->> (clojure.string/join " "))
               (clojure.string/split #"\s+")
               first
               keyword)
        v (reduce merge
            (map property-as-map (.getProperties (.getStyle rule))))]
    (assoc {} k v)))

(defn parse-css
  "Take a css file as an input and the return back the css as a hash map"
  [css]
  (let [source (-> css-file
                   slurp
                   StringReader.
                   InputSource.)
        parser (CSSOMParser.)
        stylesheet (.parseStyleSheet parser source nil nil)
        _ (println stylesheet)
        rule-list (-> stylesheet
                      .getCssRules
                      .getRules)]
    nil))

(defn adjust-list-ind-level
  [ndp]
  (when (= 0 @ilvl)
    (swap! numid (docx/restart-numbering ndp)))
  (swap! ilvl dec))

(defn same-type?
  [j-class node]
  (= (.getName j-class) (-> node (.getClass) (.getName))))

(defn get-siblings
  [node]
  (if (nil? node)
    []
    (cons node (get-siblings (.getNext node)))))

(defn get-childern
  [parent]
  (get-siblings (.getFirstChild parent)))

(defn get-grand-parent
  [node]
  (.getParent (.getParent node)))

(defn add-text
  [node p text]
  (let [parent (.getParent node)
        grand-parent (get-grand-parent node)]
    (cond
      (same-type? ListItem grand-parent) (docx/add-text-list p @numid @ilvl text)
      (same-type? StrongEmphasis parent) (docx/add-emphasis-text p :bold text)
      (same-type? Emphasis parent) (docx/add-emphasis-text p :italic text)
      :else (docx/add-text p text))))

(defn build-visitor []
  (reify Visitor

    (^void visit [this ^CustomBlock node]
     (when (same-type? TableBlock node)
       (let [table (docx/add-table maindoc)]
         (push stack table)
         (doseq [child (get-childern node)]
           (.visit (build-visitor) child))
         (pop!! stack))))

    (^void visit [this ^CustomNode node]
     (cond
       (same-type? TableRow node) (let [table (head stack)
                                        row (docx/add-table-row table)]
                                    (push stack row)
                                    (doseq [child (get-childern node)]
                                      (.visit (build-visitor) child))
                                    (pop!! stack))
       (same-type? TableCell node) (let [row (head stack)
                                         cell (docx/add-table-cell row)]
                                     (push stack cell)
                                     (doseq [child (get-childern node)]
                                       (.visit (build-visitor) child))
                                     (pop!! stack))
       :else (doseq [child (get-childern node)]
               (.visit (build-visitor) child))))

    (^void visit [this ^Document node]
     (doseq [child (get-childern node)]
       (.visit (build-visitor) child)))

    (^void visit [this ^Heading node]
     ;; TODO: Implement Heading
     (let [lvl (.getLevel node)
           p (docx/add-paragraph maindoc)]
       (docx/add-style-heading p lvl)
       (push stack p)
       (doseq [child (get-childern node)]
         (.visit (build-visitor)child))
       (pop!! stack)))

    (^void visit [this ^Paragraph node]
     (let [p (docx/add-paragraph maindoc)]
       (push stack p)
       (doseq [child (get-childern node)]
         (.visit (build-visitor) child))
       (pop!! stack)))

    (^void visit [this ^ListItem node]
     (doseq [child (get-childern node)]
       (.visit (build-visitor) child)))

    (^void visit [this ^BulletList node]
     ;; TODO: Implement Bulletlist
     (doseq [child (get-childern node)]
       (.visit (build-visitor) child)))

    (^void visit [this ^OrderedList node]
     (let [ndp (docx/add-ordered-list maindoc)]
       (push stack ndp)
       (swap! ilvl inc)
       (doseq [child (get-childern node)]
         (.visit (build-visitor) child))
       (adjust-list-ind-level ndp)
       (pop!! stack)))

    (^void visit [this ^Emphasis node]
     ;; TODO: Implelement Italic
     (doseq [child (get-childern node)]
       (.visit (build-visitor) child)))

    (^void visit [this ^StrongEmphasis node]
     ;; TODO: Implement Bold text
     (doseq [child (get-childern node)]
       (.visit (build-visitor)  child)))

    (^void visit [this ^HardLineBreak node]
     ;; TODO: Implement HardLineBreak
     (doseq [child (get-childern node)]
       (.visit (build-visitor) child)))

    (^void visit [this ^ThematicBreak node]
     ;; TODO: Implement ThematicBreak
     (doseq [child (get-childern node)]
       (.visit (build-visitor) child)))

    (^void visit [this ^SoftLineBreak node]
     ;; TODO: Implement softlinebreak
     (doseq [child (get-childern node)]
       (.visit (build-visitor) child)))

    (^void visit [this ^Text node]
     ;; TODO: Implement Text styling
     (let [p (head stack)
           text (.getLiteral node)]
       (add-text node p text)))))

(defn parse-md-common
  [s]
  (let [parser (.build (.extensions (Parser/builder) (list (TablesExtension/create))))
        tree (.parse parser s)]
    (.visit (build-visitor) tree)))

(defn -main [& args]
  (parse-md-common (slurp md-file))
  (docx/save package docx-file))
