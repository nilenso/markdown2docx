(ns markdown2docx.core
  (:require [clojure.java.io :refer [file]]
            [clojure.string :as string])
  (:import [org.docx4j.openpackaging.packages WordprocessingMLPackage]
           [com.steadystate.css.parser CSSOMParser]
           [org.w3c.dom.css CSSStyleSheet CSSRuleList CSSRule CSSStyleRule CSSStyleDeclaration]
           [org.w3c.css.sac InputSource]
           [java.io StringReader]))

;; Temp files
(def md-file "/home/sp4de/Nilenso/cooperative-agreement/template.md")
(def docx-file "/home/sp4de/Nilenso/cooperative-agreement/template.docx")
(def css-file "/home/sp4de/Nilenso/cooperative-agreement/style.css")

(defn zig []
  (let [pkg (WordprocessingMLPackage/createPackage)]
    (do (-> pkg
            (.getMainDocumentPart)
            (.addParagraphOfText "HAXED BY CHINESE!")))
    (.save pkg (file docx-file))))

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
            (map property-as-hash-map (.getProperties (.getStyle rule))))]
    (assoc {} k v)))

(defn parse-css [css]
  (let [source (-> css-file
                   slurp
                   StringReader.
                   InputSource.)
        parser (CSSOMParser.)
        stylesheet (.parseStyleSheet parser source nil nil)
        rule-list (-> stylesheet
                      .getCssRules
                      .getRules)]
    (reduce merge (map rule-as-hash-map rule-list))))

(defn -main [& args]
  ;; (println (slurp md-file))
  ;; (println (slurp css-file))
  (parse-css css-file)
  ;; TODO: take in params: markdown, css, docx
  ;; TODO: read markdown, look for css ids / classes
  ;; TODO: read in css, match classes
  ;; TODO: write to docx using corresponding docx styles for each css style
  )
