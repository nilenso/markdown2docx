(ns markdown2docx.css-test
  (:require [markdown2docx.css :as css]
            [clojure.test :refer :all]))


(deftest test-parse
  (testing "parse css should return a valid css map"
    (is (= {:.title {:text-align "center"
                     :font-style "italics"
                     :font-weight "bold"}}
           (css/parse-css ".title{
text-align : center;
font-style : italics;
font-weight : bold;}")))))
