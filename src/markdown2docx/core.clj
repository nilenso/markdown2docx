(ns markdown2docx.core
  (:require [clojure.java.io :refer [file]]
            [clojure.string :as string])
  (:use [endophile.core :only [mp to-clj html-string]])
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
        rule-list (-> stylesheet
                      .getCssRules
                      .getRules)]
    (reduce merge (map rule-as-map rule-list))))

(defn parse-md
  ""
  [md]
  (-> md
      slurp
      mp
      to-clj))

(defn -main [& args]
  (clojure.pprint/pprint (parse-css css-file))
  (clojure.pprint/pprint (parse-md md-file))
  ;; TODO: take in params: markdown, css, docx
  ;; TODO: read markdown, look for css ids / classes
  ;; TODO: read in css, match classes
  ;; TODO: write to docx using corresponding docx styles for each css style
  )
