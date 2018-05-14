package org.lasfs.ils.legacy.xlsimport

/**
 * Created by rpowell on 12/29/16.
 */
class Catalog {
    def types = []
    def locations = []
    def items = []

    def addLocationToCatalog(location) {
        if (!(location in this.locations)) {
            this.locations += location
        }
    }
    def addTypeToCatalog(type) {
        if (!(type in this.types)) {
            this.types += type
        }
    }
    def addItemToCatalog(item) {
        if (!(item.itemNumber in this.items)) {
            this.items[item.itemNumber] = item.itemTitle
        }
    }

    def print_catalog() {
        println "catalog"
        println "Types:" + this.types.size()
        this.types.sort().each { item -> println "  '${item}'" }
        println "Locations:" + this.locations.size()
        this.locations.sort().each { item -> println "  '${item}'" }
        println "Items: " + this.items.size()
    }
}