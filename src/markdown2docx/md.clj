(ns markdown2docx.md
  (:import [org.commonmark.parser Parser]
           [org.commonmark.html HtmlRenderer]
           [org.commonmark Extension]
           [org.commonmark.ext.gfm.tables TablesExtension TableBlock TableCell
            TableBody TableHead TableRow]
           [org.commonmark.node Block BlockQuote
            BulletList Code CustomBlock CustomNode Document
            Emphasis FencedCodeBlock HardLineBreak Heading
            HtmlBlock HtmlInline Image IndentedCodeBlock Link
            ListBlock ListItem Node OrderedList Paragraph
            SoftLineBreak StrongEmphasis Text ThematicBreak]))

(defn get-siblings
  [node]
  (if (nil? node)
    []
    (cons node (get-siblings (.getNext node)))))

(defn get-children
  [parent]
  (get-siblings (.getFirstChild parent)))

(defprotocol Visitor
  (visit [this]))

(extend-protocol Visitor

  Document
  (visit [this]
    {:document
     (for [child (get-children this)]
       (visit child))})

  Heading
  (visit [this]
    {:heading
     (cons {:lvl (.getLevel this)}
           (for [child (get-children this)]
             (visit child)))})

  Paragraph
  (visit [this]
    {:paragraph
     (for [child (get-children this)]
       (visit child))})

  TableBlock
  (visit [this]
    {:table-block
     (for [child (get-children this)]
       (visit child))})

  TableHead
  (visit [this]
    {:table-head
     (for [child (get-children this)]
       (visit child))})

  TableBody
  (visit [this]
    {:table-body
     (for [child (get-children this)]
       (visit child))})

  TableRow
  (visit [this]
    {:table-row
     (for [child (get-children this)]
       (visit child))})

  TableCell
  (visit [this]
    {:table-cell
     (for [child (get-children this)]
       (visit child))})

  ListItem
  (visit [this]
    {:list-item
     (for [child (get-children this)]
       (visit child))})

  BulletList
  (visit [this]
    {:bullet-list
     (for [child (get-children this)]
       (visit child))})

  OrderedList
  (visit [this]
    {:ordered-list
     (for [child (get-children this)]
       (visit child))})

  Emphasis
  (visit [this]
    {:italic
     (for [child (get-children this)]
       (visit child))})

  StrongEmphasis
  (visit [this]
    {:bold
     (for [child (get-children this)]
       (visit child))})

  HardLineBreak
  (visit [this]
    {:hard-line-break
     (for [child (get-children this)]
       (visit child))})

  ThematicBreak
  (visit [this]
    {:thematic-break
     (for [child (get-children this)]
       (visit child))})

  SoftLineBreak
  (visit [this]
    {:soft-line-break
     (for [child (get-children this)]
       (visit child))})

  Text
  (visit [this]
    {:text (.getLiteral this)}))


(defn parse
  [md-string]
  (let [extention (list (TablesExtension/create))
        parser (.build (.extensions (Parser/builder) extention))
        md-document (.parse parser md-string)]
    (visit md-document)))
