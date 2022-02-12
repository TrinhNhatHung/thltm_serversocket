package ToMau;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class TM {

	public static void main(String[] args) {
		List<List<Integer>> ds = xuly(3,20,5);
		for (List<Integer> list : ds) {
			System.out.println(list);
		}
	}

	public static List<List<Integer>> xuly(int t, int n, int m) {
		List<List<Integer>> res = new ArrayList<>();
		List<Integer> ds = new ArrayList<>();
		for (int i=1;i<=n;i++) ds.add(i);
		Collections.shuffle(ds);
		List<String> ktphong = new ArrayList<>();
		//List<String> ktgv = new ArrayList<>();
		while (t>0) {
			t--;
			List<Integer> ca = new ArrayList<>();
			boolean[] kt = new boolean[n+1];
			for (int i=1;i<=m;i++) {
				for (Integer x : ds) {
					String s=x+" "+i;
					if (!kt[x] && !ktphong.contains(s)) {
						ktphong.add(s);
						ca.add(x);
						kt[x]=true;
						break;
					}
				}
				for (Integer x : ds) {
					String s=x+" "+i;
					if (!kt[x] && !ktphong.contains(s)) {
						ktphong.add(s);
						ca.add(x);
						kt[x]=true;
						break;
					}	
				}
			}
			res.add(ca);
			Collections.shuffle(ds);
		}
		return res;
	}

}
