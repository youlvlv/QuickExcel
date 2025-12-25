package com.lizhiwei.quickExcel.entity;

import org.apache.poi.ss.usermodel.Picture;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Optional;

public class PictureMap extends HashMap<PictureMap.PictureKey, Picture> {




	public record PictureKey(Integer row, Integer column) {
	}
}
