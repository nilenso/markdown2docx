(ns markdown2docx.css
  (:import [com.steadystate.css.parser CSSOMParser]
           [org.w3c.dom.css CSSStyleSheet CSSRuleList CSSRule CSSStyleRule
            CSSStyleDeclaration]
           [org.w3c.css.sac InputSource]
           [java.io StringReader]))

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
  [css-string]
  (let [source (-> css-string
                   StringReader.
                   InputSource.)
        parser (CSSOMParser.)
        stylesheet (.parseStyleSheet parser source nil nil)
        rule-list (-> stylesheet
                      .getCssRules
                      .getRules)]
    (reduce merge (map rule-as-map rule-list))))
