package com.lizhiwei.quickExcel.entity;

import org.apache.poi.ss.usermodel.Picture;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Optional;

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
		if (super.get(key) == null) {
			return null;
		}
		if (super.get(key).size() <= index) {
			return null;
		}
		return Optional.ofNullable(super.get(key)).orElse(new ArrayList<>()).get(index);
	}
}
