package com.github.fisherman08.poik

import io.kotlintest.shouldBe
import io.kotlintest.specs.StringSpec
import java.io.File

class PoikPickerTest: StringSpec() {
    private val file = File(javaClass.getResource("/com/github/fisherman08/poik/test.xlsx").file)

    init {

        val picker = PoikPicker(file = file)

        "Poik" {
            picker.string(0, 0) shouldBe  "2"

            picker.exec { p ->

            }
        }


        picker.close()
    }
}