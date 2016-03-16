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
               ;; HACK: transforming a list into a string to split it immediately is not semantically sound
               (->> (clojure.string/join " "))
               (clojure.string/split #"\s+")
               first
               keyword)
        v (reduce merge
            (map property-as-map (.getProperties (.getStyle rule))))]
    {k v}))

(defn rule-as-map-hack-alterative [rule]
  ;; HACK: the entire thread below is kind of a hack -- it's okay because we have a really specific use case
  ;;       and we're not implementing all of the CSS spec (or even comprehensively approaching a portion of it)
  (let [k  (-> rule
               .getSelectors ;; returns SelectorListImpl
               .getSelectors ;; returns a List<Selector>
               first
               .toString
               keyword)
        v (reduce merge
            (map property-as-map (.getProperties (.getStyle rule))))]
    {k v}))

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
