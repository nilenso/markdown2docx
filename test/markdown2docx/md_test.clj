(ns markdown2docx.md-test
  (:require [clojure.test :refer :all]
            [markdown2docx.md :refer :all]))

(deftest parse-table
  (testing "parse md tables"
    (let [table-md "
| Tables        | Are           | Cool  |
| ------------- |:-------------:| -----:|
| col 3 is      | right-aligned | $1600 |
| col 2 is      | centered      |   $12 |
| zebra stripes | are neat      |    $1 |
"
          table-clj {:document
                     [{:table-block
                       [{:table-head
                         [{:table-row
                           [{:table-cell [{:text "Tables"}]}
                            {:table-cell [{:text "Are"}]}
                            {:table-cell [{:text "Cool"}]}]}]}
                        {:table-body
                         [{:table-row
                           [{:table-cell [{:text "col 3 is"}]}
                            {:table-cell [{:text "right-aligned"}]}
                            {:table-cell [{:text "$1600"}]}]}
                          {:table-row
                           [{:table-cell [{:text "col 2 is"}]}
                            {:table-cell [{:text "centered"}]}
                            {:table-cell [{:text "$12"}]}]}
                          {:table-row
                           [{:table-cell [{:text "zebra stripes"}]}
                            {:table-cell [{:text "are neat"}]}
                            {:table-cell [{:text "$1"}]}]}]}]}]}]
      (is (= table-clj (parse table-md))))))

(deftest parse-heading
  (let [get-lvl #(-> % :document first :heading first :lvl)]

    (testing "parse returns a representative map"
      (is (= {:document [{:heading [{:lvl 1}
                                    {:text "Heading"}]}]}
             (parse "# Heading"))))

    (testing "parse correctly identifies a level 1 heading"
      (is (= 1 (get-lvl (parse "# Heading")))))

    ;; TODO: extract the remaining testing blocks
    (testing "GIVEN <some-precondition> WHEN <something-happens> THEN <should-be-result>"
      (let [h2 "## Smaller Heading"
            h3 "### Even smaller Heading ?"
            h4 "#### This is even smaller"
            h5 "##### A little bit smaller"
            h6 "###### smallest heading, hopefully"]
        (is (= 2 (get-lvl (parse h2))))
        (is (= 3 (get-lvl (parse h3))))
        (is (= 4 (get-lvl (parse h4))))
        (is (= 5 (get-lvl (parse h5))))
        (is (= 6 (get-lvl (parse h6))))))))

(deftest parse-emphasis
  (testing "parse bold"
    (let [obi-wan "**In my experience there is no such thing as luck.**"
          obi-wan-clj {:document
                       [{:paragraph
                         [{:bold
                           [{:text "In my experience there is no such thing as luck."}]}]}]}]
      (is (= obi-wan-clj (parse obi-wan)))))

  (testing "parse italics"
    (let [yoda "*When nine hundred years old you reach, look as good you will not.*"
          yoda-clj {:document
                       [{:paragraph
                         [{:italic
                           [{:text "When nine hundred years old you reach, look as good you will not."}]}]}]}]
      (is (= yoda-clj (parse yoda)))))

  (testing "parse both bold and italics together"
    (let [yoda "**Do**. Or **do not**. There is no *try*"
          yoda-clj {:document
                    [{:paragraph
                      [{:bold [{:text "Do"}]}
                       {:text ". Or "}
                       {:bold [{:text "do not"}]}
                       {:text ". There is no "}
                       {:italic [{:text "try"}]}]}]}]
      (is (= yoda-clj (parse yoda))))))

(deftest parse-ordered-lists
  (testing "test ordered list"
    (let [quotes "1. May the Force be with you
  2. I find your lack of faith disturbing.
  3. I've got a very bad feeling about this.
  4. Never tell me the odds!
  5. Truly wonderful, the mind of a child is."
          quotes-clj {:document
                      [{:ordered-list
                        [{:list-item [{:paragraph [{:text "May the Force be with you"}]}]}
                         {:list-item [{:paragraph [{:text "I find your lack of faith disturbing."}]}]}
                         {:list-item [{:paragraph [{:text "I've got a very bad feeling about this."}]}]}
                         {:list-item [{:paragraph [{:text "Never tell me the odds!"}]}]}
                         {:list-item [{:paragraph [{:text "Truly wonderful, the mind of a child is."}]}]}]}]}]
      (is (= quotes-clj (parse quotes)))))

  (testing "test nested ordered list"
    (let [nested-list "1. First
     2. Second
        3. Third"
          nested-list-clj {:document
                           [{:ordered-list
                             [{:list-item [{:paragraph [{:text "First"}]}
                                {:ordered-list
                                 [{:list-item [{:paragraph [{:text "Second"}]}
                                    {:ordered-list
                                     [{:list-item [{:paragraph [{:text "Third"}]}]}]}]}]}]}]}]}]
      (is (= nested-list-clj (parse nested-list))))))
