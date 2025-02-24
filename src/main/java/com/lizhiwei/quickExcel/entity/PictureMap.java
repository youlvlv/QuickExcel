package com.lizhiwei.quickExcel.entity;

import org.apache.poi.ss.usermodel.Picture;

import java.util.HashMap;
import java.util.List;

public class PictureMap extends HashMap<Integer, List<Picture>> {

	@Override
	public List<Picture> put(Integer key, List<Picture> value) {
		return super.put(key, value);
	}

	public synchronized List<Picture> put(Integer key, Picture value) {
		List<Picture> pictures = super.get(key);
		if (pictures == null) {
			pictures = new java.util.ArrayList<>();
		}
		pictures.add(value);
		return super.put(key, pictures);
	}

	public Picture get(Integer key, Integer index) {
		return super.get(key).get(index);
	}
}
