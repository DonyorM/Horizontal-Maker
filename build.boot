(set-env!
 :resource-paths #{"src" "tests"}
 :dependencies '[[org.clojure/clojure "1.9.0"]
                 [seesaw "1.4.5"]
                 [dk.ative/docjure "1.12.0"]])

(task-options!
 pom {:project 'horizontal-maker/horizontal-maker
      :version "1.0"
      :description "Helper for generating horizontal charts for SBS and BCC"
      :developers {"Daniel Manila" "daniel.develop@manilas.net"}
      :license {"MIT" "https://opensource.org/licenses/MIT"}})

(require 'horizontal-maker.start)
(deftask run []
  (with-pass-thru _
    (horizontal-maker.start/-main)))

(deftask build []
  (set-env! :resource-paths #{"src"})
  (comp
   (pom)
   (aot :namespace #{'horizontal-maker.start})
   (uber)
   (jar :main 'horizontal-maker.start :file "horizontal-maker.jar")
   (sift :include #{#"horizontal-maker.jar"})))

