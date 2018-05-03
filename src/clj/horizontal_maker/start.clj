(ns horizontal-maker.start
  (:gen-class))

(defn -main [& args]
  (require 'horizontal-maker.gui)
  ((resolve 'horizontal-maker.gui/start-gui)))
