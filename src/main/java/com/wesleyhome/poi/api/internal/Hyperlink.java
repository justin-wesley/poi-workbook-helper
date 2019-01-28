package com.wesleyhome.poi.api.internal;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
@AllArgsConstructor
public class Hyperlink {
    private String url;
    private String linkText;

    public String toString(){
        return linkText;
    }
}
